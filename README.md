# Code-archive-VBA-p1

Ce projet illustre mes premiers pas dans le développement avec VBA, réalisés dans le but de digitaliser et réorganiser l’administration d’une boutique.

## Partie 1 : Programme de comptabilité

Cette première étape m’a permis de me familiariser avec le langage VBA. Vers 2007–2008, les solutions logicielles étaient bien plus coûteuses et la boutique ne pouvait pas se permettre un programme "tout-en-un". J’ai donc commencé à créer mes propres outils sur Excel pour répondre aux besoins du quotidien.

À l’époque, j’utilisais l’éditeur de code d’Excel : chaque action dans le tableur était traduite automatiquement en code.
 Je décortiquais ces transcriptions pour comprendre quelle instruction correspondait à chaque manipulation, puis je les combinais pour construire pas à pas un petit programme comptable.

De fil en aiguille, je découvrais les spécificités de VBA jusqu’à pouvoir écrire directement le code de manière autonome.
 Le programme prenait forme, les séances de débogage s’enchaînaient. Bien sûr, j’étais à des années-lumière d’un vrai logiciel comptable, mais il m’aidait déjà à automatiser une petite tâche.

Le gain de temps réel était minime par rapport au temps investi… mais j’avais acquis mes premières bases solides en VBA.

---

## Difficultés rencontrées

À cette époque, l’accès à l’information était bien plus limité qu’aujourd’hui. Les ressources disponibles en ligne étaient moins nombreuses, ce qui rendait l’apprentissage et le développement plus laborieux.

---

## Explications

Le fichier **Fiduciaire_.bas** contient plusieurs procédures (`Sub`) destinées à être associées à des boutons personnalisés dans la barre d’outils (fonction disponible uniquement sur la version PC, absente de Microsoft Office 2011).

### Liste des procédures
- `Sub Bouton_Nouveau_Fichier_Comptabilité()`
- `Sub Bouton_Transfert_Comptabilité_à_Séparations()`
- `Sub Bouton_Nouveau_Fichier_Mensuel()`
- `Sub Bouton_Nouveau_Fichier_Liste_et_Séparations()`
- `Sub Bouton_Transfert_Liste_à_Séparations()`

### Ordre d’exécution conseillé

1. **`Bouton_Nouveau_Fichier_Liste_et_Séparations()`**  
   Crée le fichier **Comptes.xlsx** (⚠️ à enregistrer manuellement sous ce nom).  
   Ce fichier joue le rôle d’un plan comptable personnalisable :  
   - `C` pour les charges  
   - `D` pour les débiteurs  
   - `F` pour les fournisseurs  
   - `L` pour les liquidités  

   La feuille *Séparations* génère des intercalaires en fonction des comptes listés dans la feuille *Liste*.  
   Ces intercalaires pouvaient ensuite être imprimés et utilisés pour classer physiquement les documents dans des classeurs (souvenir de l’époque du **100% papier** 📂).

2. **`Bouton_Transfert_Liste_à_Séparations()`**  
   Transfère automatiquement les comptes inscrits dans la feuille *Liste* vers la feuille *Séparations*.  

3. **`Bouton_Nouveau_Fichier_Comptabilité()`**  
   Crée le fichier **Comptabilité.xlsx**, enregistré automatiquement dans le dossier *Fiduciaire*.  
   Pour chaque compte défini dans la feuille *Liste*, la macro crée **12 fiches mensuelles** permettant de saisir :  
   - le détail des factures  
   - le nom de l’entreprise  
   - la date de paiement  
   - le montant brut  
   - la TVA  

   Chaque feuille calcule automatiquement le total des frais et de la TVA déductible.  

4. **`Bouton_Nouveau_Fichier_Mensuel()`**  
   Crée le fichier **Mensuel.xlsx**, qui génère un résumé mensuel des totaux de tous les comptes.  
   (Utile car les feuilles du fichier *Comptabilité.xlsx* deviennent rapidement volumineuses.)

5. **`Bouton_Transfert_Comptabilité_à_Séparations()`**  
   Transfère les noms des comptes et leurs totaux mensuels respectifs de *Comptabilité.xlsx* à *Mensuel.xlsx*.   

### Autres fichiers
Les fichiers **XLSX** fournis dans le repository sont des aperçus visuels des résultats générés par les macros contenues dans le fichier **Fiduciaire_.bas**.

