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
2. Démarrez un diaporama dans PowerPoint.
3. Prononcez l'une des commandes vocales ci-dessus.

L'application reste dans la zone de notification Windows. Un clic droit sur son
icône permet de mettre l'écoute en pause, de tester la commande PowerPoint ou de
choisir **Français / English**. Le choix de langue est mémorisé. Un double-clic
active ou suspend aussi l'écoute.

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

## Licence

Ce projet est distribué sous [licence MIT](LICENSE).
