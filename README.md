# automatisation-tabac-arcpy

Ce script permet d'automatiser la procédure d'une étude géomarketing sur l'implantation de débits de tabac selon le cadre législatif.

Quatres grands axes de la méthodologie sont automatique :
  - Le géocodage des adresses 
  - Le calcul de la zone de desserte
  - Le calcul des distances 
  - Le rassemblement de toutes les données .xls dans un seul tableau excel
  
Pour le bon fonctionnement du script, on dispose de deux fichiers xls "target.xls" et "tabac.xls" :
  - Le fichier "target.xls" représente le lieu d'implantation du tabac
  - Le fichier "tabac.xls" représente les tabacs existants dans le département du lieu d'implantation
  
Saisissez votre adresses dans le champs "adresse" en fonction de vos besoins. 

Ensuite, chercher vos données communes sous format shapefile et les insérer dans le répertoire "communes"
(un exemple est à votre disposition dans le dossier "communes)

Avant de lancer le script, changer les paramètre de la fonction "desserte_et_distance()" dans main.py par vos identifiants arcgis.

Besoin d'aide ? Contactez-moi à l'addresse : laurentwu123@gmail.com
©Laurent WU
