## Extractions à partir de l'ERP
Extractions programmées et envoyées par mail du lundi au vendredi à heure fixe

## Power Automate Desktop : Suppression des anciennes extractions et remplacement par les nouvelles
- Flux principal à partir duquel sont appelés les sous-flux :
<img width="1349" height="310" alt="image" src="https://github.com/user-attachments/assets/78c38971-2401-474a-baf7-fa0bc56a35f8" />


- Extrait d'un sous-flux de nettoyage des anciens fichiers / remplacement par les nouveaux :
<img width="1568" height="834" alt="image" src="https://github.com/user-attachments/assets/756a5377-d35f-4850-ab2b-4bb4999c3fe0" />

## Power Automate Desktop : Manipulation des fichiers (filtrage et suppression des colonnes inutiles)
- Extrait d'un sous-flux de traitement :
<img width="1539" height="854" alt="image" src="https://github.com/user-attachments/assets/a9514195-5522-4d42-9f89-848ce4a2648f" />

## Power Automate Desktop : Résultat final
Le workflow s'exécute en arrière-plan, en quelques minutes les extractions sont remplacées et rangées, prêtes pour l'import via VBA.

- Voici l'arborescence de dossiers utilisée pour le rangement des extractions sur une région :
<img width="707" height="556" alt="image" src="https://github.com/user-attachments/assets/745197fa-81a7-441a-9ac8-a0b98066805f" />


L'appel VBA permet ensuite d'aller piocher les plages désirées dans chaque fichier exporté pour les coller dans les plages spécifiées du reporting.
