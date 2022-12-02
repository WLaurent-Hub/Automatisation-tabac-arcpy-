import arcpy
import os
from pathlib import Path
import sys
import time
import json
import glob
import shutil
import pandas as pd

arcpy.CreateFileGDB_management(str(Path(__file__).parent), "geodatabase.gdb")  # création d'une géodatabase
os.mkdir(str(Path(__file__).parent)+"/data")

#---------- geocodage --------#
def geocoding(table, field, out):
    
    database = str(Path(__file__).parent) + "/" + table +"$" # selection de la database débit de tabac
    standard_geocod = "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/ArcGIS World Geocoding Service" # geocode universel
    output = str(Path(__file__).parent)+ "/geodatabase.gdb/" + out # sortie du fichier en geodatabase
    arcpy.geocoding.GeocodeAddresses(database, standard_geocod, field, output) # processus de géocoding
    
    directory_out = str(Path(__file__).parent) + "/" + out
    os.mkdir(directory_out) # création d'un dossier

    arcpy.FeatureClassToShapefile_conversion(output, directory_out) # conversion geodatabase en shapefile

#------ Zone de dessert / Ressource la plus proche --------#
def desserte_et_distance(identifiant,mdp):
    
    global output_service_areas, incidents
    
    username = identifiant; # votre identifiants arcgis (saisir dans les paramètres)
    password = mdp # votre mot de passe arcgis (saisir dans les paramètres)
    sa_service = "https://logistics.arcgis.com/arcgis/services;World/ServiceAreas;{0};{1}".format(username, password)
    cf_service = "https://logistics.arcgis.com/arcgis/services;World/ClosestFacility;{0};{1}".format(username, password)

    # import la boite à outil
    arcpy.ImportToolbox(sa_service)
    arcpy.ImportToolbox(cf_service)

    os.mkdir(str(Path(__file__).parent) + "/desserte_et_distance") # création d'un dossier zone de desserte
    facilities = str(Path(__file__).parent)+"/geodatabase.gdb/implantation" # table d'entrée 
    incidents = str(Path(__file__).parent)+"/geodatabase.gdb/tabac"  
    output_service_areas = str(Path(__file__).parent) + "/desserte_et_distance/Desserte" # table de sortie
    output_routes = str(Path(__file__).parent)+"/desserte_et_distance/Routes"

    # paramètre distance à pied 
    portal_url = "https://www.arcgis.com"
    arcpy.SignInToPortal(portal_url, username, password) # connection à arcgis
    travel_mode_list = arcpy.na.GetTravelModes(portal_url) # dictionnaire d'objets de mode de déplacement
    walking_distance = travel_mode_list["Walking Distance"] # paramètre distance à pied
    walking_distance_str = str(walking_distance) # conversion en string
  
    # génère la zone de chalandise
    result_desserte = arcpy.ServiceAreas.GenerateServiceAreas(facilities, "1,5", "Kilometers", Travel_Mode=walking_distance_str)
    result_closestFacility = arcpy.ClosestFacility.FindClosestFacilities(incidents, facilities, "Kilometers", Travel_Mode=walking_distance_str)
    #vérifie le processus chaque seconde
    while result_desserte.status < 4:
        time.sleep(1)

    # imprime les messsages d'erreur du processus 
    result_severity = result_desserte.maxSeverity
    if result_severity == 2:
        arcpy.AddError("Une erreur s'est produite lors de l'exécution") # message d'erreur personnalisé
        arcpy.AddError(result_desserte.getMessages(2), result_closestFacility.getMessages(2) ) # gravité de l'erreur
        sys.exit(2)
    elif result_severity == 1:
        arcpy.AddWarning("Un avertissement s'est produit lors de l'exécution ") # message d'avertissement personnalisé
        arcpy.AddWarning(result_desserte.getMessages(1), result_closestFacility.getMessages(1)) # gravité de l'erreur

    # génère un shape du processus en sortie
    result_desserte.getOutput(0).save(output_service_areas)
    result_closestFacility.getOutput(0).save(output_routes)

#--- Selection par attribut + conversion xls ---#
def convertExcel(input,output, excel):
    
    # selection par attribut -- intersection entre deux couches
    debit_in_desserte = arcpy.SelectLayerByLocation_management(in_layer = str(Path(__file__).parent)+"/"+ input, 
                                                               overlap_type='INTERSECT',
                                                               select_features = str(Path(__file__).parent) + "/" + output, 
                                                               selection_type='NEW_SELECTION')

    # conversion des entités sélectionnés en tableau excel
    arcpy.conversion.TableToExcel(debit_in_desserte, str(Path(__file__).parent)+"/" + excel)
        
#--- Regroupe toutes les données (xls) dans un seul tableau xls (data.xls) ---#        
def mergeXls():
    excel_files = glob.glob(str(Path(__file__).parent) + "/data/*.xls")
    writer = pd.ExcelWriter(str(Path(__file__).parent)+ "/data.xls") # enregistrement dataframe dans une feuille Excel

    for excel_file in excel_files: 
        sheet = os.path.basename(excel_file)
        sheet = sheet.split(".")[0] # extraction du filename sans extension
        df1 = pd.read_excel(excel_file) # lit les fichiers xls dans un dataframe
        df1.fillna(value="N/A", inplace=True) # remplacement des données vides
        df1.to_excel(writer,sheet_name=sheet,index=False) # exporte le dataframe en excel

    writer.save()
    writer.close()
    
    for xls in glob.glob(str(Path(__file__).parent) + "/*.xls"):
        shutil.move(xls, str(Path(__file__).parent)+ "/data")
    
geocoding("tabac.xls/Feuil1", "\'Address or Place\' adresse", "tabac")
geocoding("target.xls/Feuil1", "\'Address or Place\' adresse", "implantation")    
desserte_et_distance("votre_id","mdp") # mettre ses identifiants dans les paramètres
convertExcel("geodatabase.gdb/tabac","desserte_et_distance/Desserte","tabac_desserte.xls")
convertExcel("communes/communes.shp","desserte_et_distance/Desserte","communes_desserte.xls")
arcpy.conversion.TableToExcel(str(Path(__file__).parent) + "/desserte_et_distance/Routes.shp", str(Path(__file__).parent)+"/Routes.xls")
mergeXls()
