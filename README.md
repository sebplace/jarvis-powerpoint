# Jarvis PowerPoint

Petite application Windows qui pilote le diaporama PowerPoint lorsque vous dites :

- **« Jarvis, suivant ! »** pour avancer ;
- **« Jarvis, précédent ! »** pour revenir en arrière ;
- **« Jarvis, va au slide X ! »** pour atteindre directement la diapositive X
  (de 1 à 999) ;
- **« Jarvis, cherche budget »** ou **« Jarvis, va au slide sur la sécurité »**
  pour rechercher un titre ou du texte visible.

Le mode anglais, sélectionnable depuis l'icône de notification, accepte :

- **“Jarvis, next!”** pour avancer ;
- **“Jarvis, previous!”** pour revenir en arrière ;
- **“Jarvis, go to slide X!”** pour atteindre directement la diapositive X
  (de 1 à 999) ;
- **“Jarvis, search for budget”** ou **“Jarvis, go to the slide about security”**
  pour rechercher un titre ou du texte visible.

## Télécharger

Téléchargez `JarvisPowerPoint.exe` depuis la
[dernière version publiée](https://github.com/sebplace/jarvis-powerpoint/releases/latest).
Aucune installation n'est nécessaire.

Placez l'exécutable dans un dossier où vous souhaitez le conserver avant de le
lancer. Au premier lancement, Jarvis propose de créer un raccourci dans votre
menu Démarrer, sans droits administrateur. Acceptez pour retrouver ensuite
l'application en appuyant sur **Windows** et en tapant **Jarvis PowerPoint**.

Si vous refusez, la proposition ne se répète pas aux lancements suivants. Vous
pouvez créer le raccourci plus tard avec un clic droit sur l'icône Jarvis près de
l'horloge, puis **Créer / actualiser le raccourci Démarrer...**. Un raccourci déjà
présent est conservé, sans nouvelle demande au démarrage.

L'application reste portable : le raccourci pointe vers l'exécutable, il ne le
copie pas et n'active pas le démarrage automatique avec Windows. Si vous déplacez
l'exécutable, relancez-le depuis son nouveau dossier et actualisez le raccourci
depuis ce même menu. Si vous supprimez Jarvis, supprimez aussi son raccourci.

Windows peut afficher un avertissement SmartScreen, car l'exécutable n'est pas
signé numériquement.

## Utilisation

1. Lancez `JarvisPowerPoint.exe`.
2. Dans l'assistant initial, choisissez la langue et le microphone. Le bouton de
   test permet de prononcer **« Jarvis test »** sans agir sur PowerPoint. Le test
   reste facultatif ; enregistrez les réglages pour terminer.
3. Démarrez un seul diaporama dans PowerPoint.
4. Prononcez l'une des commandes vocales ci-dessus.

L'application reste dans la zone de notification Windows. Un clic droit sur son
icône permet de mettre l'écoute en pause, de tester la commande PowerPoint ou de
choisir **Français / English**. Le choix de langue est mémorisé. Un double-clic
active ou suspend aussi l'écoute.

Les réglages sont accessibles à tout moment via **Langue et microphone...**.
La sélection du microphone ne modifie pas le périphérique par défaut des autres
applications Windows. Si un microphone choisi n'est plus disponible, Jarvis le
signale : sélectionnez-le à nouveau ou choisissez le microphone Windows par défaut.
Mettre l'écoute en pause libère le microphone utilisé par Jarvis.

## Rechercher, répondre, puis reprendre

| Français | English | Action |
|---|---|---|
| Jarvis, reprends la présentation | Jarvis, resume the presentation | Revenir au point mémorisé |
| Jarvis, autre résultat | Jarvis, another result | Parcourir les résultats de la dernière recherche |
| Jarvis, mode questions | Jarvis, questions mode | Mémoriser la position avant les questions-réponses |
| Jarvis, écran noir | Jarvis, black screen | Masquer temporairement les slides |
| Jarvis, affiche | Jarvis, restore slides | Réafficher le diaporama |
| Jarvis, raccourci la démo | Jarvis, shortcut the demo | Ouvrir un alias personnel |
| Jarvis, démarre la répétition | Jarvis, start rehearsal | Démarrer le chronométrage |
| Jarvis, arrête la répétition | Jarvis, stop rehearsal | Arrêter et conserver le bilan en mémoire |

Un saut par recherche, numéro ou alias mémorise le point de départ. Les sauts
suivants ne l'écrasent pas : dites **« Jarvis, reprends »** pour revenir, puis
continuer normalement. Le mode questions permet de poser ce repère avant même
de naviguer. La reprise réussie efface le repère et les résultats de recherche.
Le repère appartient au diaporama courant, pas à toutes les présentations.

**Autre résultat** parcourt les correspondances dans l'ordre du classement,
puis revient au premier résultat. Les références utilisent l'identifiant des
diapositives pour résister à leur réorganisation. Une diapositive supprimée ne
doit pas vous envoyer sur une autre par erreur. Une recherche sans résultat ne
déplace pas le diaporama.

## Outils présentateur

Clic droit sur l'icône Jarvis, puis **Outils présentateur...** :

- **Présentateur** : état du microphone, niveau sonore, dernière commande
  comprise et résultat de son exécution ; boutons de reprise, questions,
  résultats et écran noir.
- **Alias de slides** : associez un nom au slide courant d'une présentation
  enregistrée. Le même nom met à jour l'association ; sélectionnez un alias
  pour le supprimer ou le tester. Il est propre à cette présentation.
- **Répétition** : chronométrage cumulé par slide, budget par défaut de
  90 secondes, dépassements affichés en rouge. Sélectionnez une ligne pour
  lui appliquer un budget différent. Les retours sur un slide cumulent le temps.

L'indicateur compact est **désactivé par défaut**. Activez-le dans le panneau et
choisissez explicitement l'écran sur lequel l'afficher. Il ne vole pas le focus
au diaporama. Fermer le panneau le masque sans quitter Jarvis.

**Attention au partage d'écran :** le panneau et l'indicateur sont des fenêtres
Windows visibles, pas une zone protégée. Placez-les sur votre écran présentateur
et partagez uniquement la fenêtre du diaporama. Ils seront visibles si vous
partagez tout l'écran qui les contient.

## Répétition et données locales

Le chronométrage suit également les changements de slide effectués au clavier
ou à la souris (échantillonnage toutes les 500 ms). Le temps passé en écran noir
est exclu. La fin ou le changement de diaporama arrête la répétition. Une
nouvelle répétition remplace le bilan précédent.

Le bilan reste en mémoire jusqu'à la fermeture de Jarvis. Le bouton **Exporter
CSV...** permet de l'enregistrer dans un fichier choisi : titre de présentation,
numéro et titre de slide, temps, budget et dépassement. Aucun audio n'est enregistré.

Les alias sont enregistrés sous `%LOCALAPPDATA%\JarvisPowerPoint`, sans modifier
le fichier PowerPoint. Leur association dépend de l'emplacement du fichier :
après un déplacement ou un renommage de la présentation, recréez ses alias.
La langue et le microphone sont mémorisés dans
`HKEY_CURRENT_USER\Software\JarvisPowerPoint`.

La reconnaissance vocale reste entièrement locale. Aucune donnée audio n'est
envoyée vers un service en ligne.

La recherche ignore les majuscules et les accents. Les titres sont prioritaires
sur le reste du contenu. Le texte intégré dans une image n'est pas indexé.

## Prérequis

- Windows avec PowerPoint installé ;
- un microphone configuré ;
- le module Windows **Reconnaissance vocale - Français (France)**.

Le mode anglais nécessite également le module **Speech recognition - English
(United States)**.

Si le module manque : ouvrez **Paramètres > Heure et langue > Langue et région**,
puis installez les options vocales de **Français (France)**.

## Recompiler

Dans Windows PowerShell :

```powershell
.\build.ps1
```

Le script utilise le compilateur .NET Framework déjà inclus dans Windows et ne
nécessite aucun SDK supplémentaire.

Pour compiler sans remplacer l'exécutable en cours d'utilisation :

```powershell
.\build.ps1 -OutputDirectory .\bin\candidate
```

## Licence

Ce projet est distribué sous [licence MIT](LICENSE).
