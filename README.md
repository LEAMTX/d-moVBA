# 📊 VBA - Macros Diverses pour Excel

Ce projet regroupe un ensemble de **macros VBA utiles** pour manipuler des données dans Excel. Chaque macro est **liée à un bouton** dans le fichier `.xlsm` pour une utilisation facile, même sans connaissance en programmation.

## 📁 Fichier principal

- **vbademo1.xlsm** : contient toutes les macros décrites ci-dessous, avec des **boutons interactifs** pour les exécuter.

---

## 🔧 Liste des macros disponibles

### 1. `Somme()`
> Affiche une **InputBox en boucle** pour saisir des entiers. La boucle continue tant que l’utilisateur clique sur "Oui". À la fin, affiche la **somme totale**.

---

### 2. `Sommelig()`
> Calcule la **somme d’une ligne** entière, à partir de la cellule active, **vers la droite** jusqu’à la première cellule vide.

---

### 3. `sommecol()`
> Calcule la **somme d’une colonne**, à partir de la cellule active, **vers le bas** jusqu’à la première cellule vide.

---

### 4. `sommediag()`
> Calcule la **somme d’une diagonale descendante** (ligne+1, colonne+1 à chaque étape), depuis la cellule active, jusqu’à une cellule vide.

---

### 5. `modulo()`
> Demande un **nombre entier** et affiche **tous ses diviseurs** :
- Affichés horizontalement à partir de la cellule B1
- Liste également affichée dans une **MsgBox**

---

### 6. `maxmin()`
> Analyse une **colonne** à partir de la cellule active et calcule :
- Le **maximum**
- Le **minimum**
- L’**amplitude** (différence entre les deux)

Résultats affichés dans les colonnes **G (min), H (max), I (amplitude)** à la première ligne vide.

---

### 7. `lettre()`
> Compte le **nombre d’occurrences d’une lettre** saisie par l'utilisateur dans toute la **colonne F**. Résultat affiché dans une MsgBox.

---

### 8. `annes()`
> Demande une **année** et indique si elle est **bissextile** ou non, selon les règles du calendrier grégorien.

---

### 9. `couleurArcEnCiel()`
> Applique des **couleurs différentes** à trois colonnes :
- **N** : thème Accent1 (nuance sombre)
- **M** : vert/turquoise (code 5287936)
- **O** : jaune (code 65535)

---

## 🖱️ Utilisation

Chaque macro est **liée à un bouton** dans la feuille Excel, ce qui permet de les exécuter facilement.  
Il suffit de **cliquer sur le bouton correspondant** pour lancer l’action souhaitée.

---

## 📌 Avertissements

- Le fichier doit être **activé avec les macros** pour fonctionner correctement.
- Ne pas supprimer ou renommer les boutons sans adapter le code VBA associé.

---

## 👩‍💻 Auteur

- Projet développé par **Léa**, dans un but pédagogique ou pour automatiser certaines tâches Excel.

---

