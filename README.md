# Jarvis PowerPoint

Petite application Windows qui pilote le diaporama PowerPoint lorsque vous dites :

- **« Jarvis, suivant ! »** pour avancer ;
- **« Jarvis, précédent ! »** pour revenir en arrière.

Le mode anglais, sélectionnable depuis l'icône de notification, accepte :

- **“Jarvis, next!”** pour avancer ;
- **“Jarvis, previous!”** pour revenir en arrière.

## Télécharger

Téléchargez `JarvisPowerPoint.exe` depuis la
[dernière version publiée](https://github.com/sebplace/jarvis-powerpoint/releases/latest).
Aucune installation n'est nécessaire.

Windows peut afficher un avertissement SmartScreen, car l'exécutable n'est pas
signé numériquement.

## Utilisation

1. Lancez `JarvisPowerPoint.exe`.
2. Démarrez un diaporama dans PowerPoint.
3. Dites « Jarvis, suivant ! » ou « Jarvis, précédent ! ».

L'application reste dans la zone de notification Windows. Un clic droit sur son
icône permet de mettre l'écoute en pause, de tester la commande PowerPoint ou de
choisir **Français / English**. Le choix de langue est mémorisé. Un double-clic
active ou suspend aussi l'écoute.

La reconnaissance vocale reste entièrement locale. Aucune donnée audio n'est
envoyée vers un service en ligne.

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
