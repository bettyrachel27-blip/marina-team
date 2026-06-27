Marina Team - V6 import Excel robuste

Corrections principales :
- lecture des fichiers .xlsx, .xls, .xlsb et .csv ;
- scan de tous les onglets planning, même si le nom de l'onglet n'a pas exactement le même format ;
- détection des semaines via les dates dans le tableau ;
- correction du parser : la colonne B peut contenir le poste, les prénoms sont maintenant cherchés dans les blocs de jours ;
- chaque jour est lu par groupe de 5 colonnes : prénom / début 1 / fin 1 / début 2 / fin 2 ;
- les anciennes semaines sont conservées et les semaines réimportées sont remplacées ;
- les semaines sont triées de la plus récente à la plus ancienne.

Déploiement Render :
1. Dézipper ce dossier.
2. Envoyer tous les fichiers sur GitHub en remplaçant les anciens.
3. Sur Render : Manual Deploy > Deploy latest commit.
4. Se connecter avec Rachel : identifiant rachel / mot de passe 2802.
5. Importer le fichier Excel depuis l'onglet Plannings.
