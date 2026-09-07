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

Choisissez **une seule option** depuis la
[dernière version publiée](https://github.com/sebplace/jarvis-powerpoint/releases/latest).

| Téléchargement | Utilisation |
|---|---|
| `JarvisPowerPoint.exe` | Application portable autonome : lancez-la directement, sans installation. |
| `JarvisPowerPoint-1.5.0-portable.zip` | Le même EXE portable avec sa licence : décompressez puis lancez l'EXE. |
| `JarvisPowerPoint-1.5.0-x86.msi` | Alternative d'installation pour tous les utilisateurs, destinée au déploiement informatique. |

**Le MSI est facultatif : il ne remplace ni l'EXE ni l'archive portable.**
Il ne convertit pas une copie portable existante et ne la supprime pas.
Les trois téléchargements fournissent la même application. La mise à jour
intégrée des copies portables continue de télécharger l'EXE, jamais le MSI ;
une installation MSI est mise à jour séparément par le service informatique.
Le MSI ne contourne ni les droits administrateur ni les règles de sécurité
de l'organisation.

Les fichiers v1.5 sont **non signés**. Le MSI est fourni pour évaluation :
validation ICE et qualification d'installation en environnement IT restent
nécessaires avant un déploiement en production.

### Utilisation portable

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
« Portable » signifie ici sans installateur : les préférences et profils
restent dans le registre utilisateur et `%LOCALAPPDATA%\JarvisPowerPoint`,
pas à côté de l'EXE.

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

Les formulations **« va à la diapo septante et un »** et
**« va à la diapositive nonante-neuf »** sont aussi reconnues après « Jarvis ».
Les nombres belges et français coexistent de 1 à 999, avec le moteur `fr-FR`.

Après une interruption ou une veille, un microphone choisi explicitement est
réessayé avec des délais de 1, 2, 5 puis 10 secondes, sans basculer vers un autre.
La pause, les réglages ouverts et la fermeture bloquent ces reprises.
Le **microphone Windows par défaut** nécessite une reprise explicite
(pause/reprise ou réglages), car Windows peut avoir changé le périphérique.
Après épuisement des essais, reconnectez le micro puis relancez l'écoute.
Si le pilote ou le moteur vocal ne peut pas être libéré correctement, Jarvis
bloque toute nouvelle écoute, les tests et la validation des réglages, même
après une pause ou une veille. Annulez l'assistant s'il est ouvert, quittez
complètement Jarvis, vérifiez le périphérique, puis redémarrez manuellement.
Les commandes manuelles du présentateur restent disponibles.
Les états d'attente, de reprise et d'échec apparaissent dans le panneau.

Les seuils de confiance sont plus prudents pour une recherche ou un alias
(0,85) que pour une commande fixe (0,75). Ce sont des scores heuristiques du
moteur vocal, pas des probabilités ni une garantie de reconnaissance.

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

## Recherche prudente

L'option **Confirmer les résultats ambigus avant de changer de slide** est activée
par défaut. Si plusieurs slides obtiennent des scores proches, Jarvis présente
une sélection de cinq résultats au maximum. Choisissez-en un et confirmez ;
**Annuler** laisse le diaporama et le point de reprise inchangés.

Un résultat nettement supérieur reste accessible immédiatement. Le choix prudent
se mémorise dans les réglages ; désactivez-le pour retrouver la recherche directe.
La navigation vocale et les raccourcis sont suspendus tant qu'un dialogue est
ouvert. Une sélection devenue obsolète après un changement de diaporama est refusée.
Le mode en ligne de commande `--search` reste direct et non interactif.

La fenêtre de choix n'est pas cachée au partage d'écran : partagez uniquement
la fenêtre PowerPoint, plutôt que tout votre bureau.

## Recherche progressive et index en mémoire

Les recherches de l'interface affichent une progression annulable avec
**Échap** ou **Annuler**, y compris lorsque la confirmation prudente est désactivée.
La lecture initiale est découpée sur le thread Windows de l'interface, ainsi que
le contrôle avant navigation des présentations enregistrées et inchangées.
Un appel PowerPoint individuel ne peut pas être interrompu : un slide très
complexe ou un PowerPoint occupé peut encore retarder l'annulation.

L'index normalisé est conservé uniquement en mémoire pour une présentation
enregistrée et inchangée. Les contrôles portent sur la session, le fichier,
le nombre et l'ordre des identifiants des slides. **Les présentations non
enregistrées ou modifiées restent recherchables**, sans enregistrement automatique.
Sans version de fichier fiable, Jarvis relit tout le contenu avant d'afficher
les résultats, puis avant de naviguer. Ce contrôle final est effectué sans pause
entre les étapes et peut bloquer temporairement l'interface ; l'annulation
n'est pas possible pendant ce contrôle.
La progression le signale avant de commencer. Enregistrer la présentation
améliore donc la réactivité des recherches suivantes, mais n'est pas obligatoire.
Le mode direct `--search` reste compatible avec les présentations non enregistrées,
sans fenêtre de progression ni confirmation.
**Actualiser l'index de recherche** dans le menu efface le cache ;
la prochaine recherche le reconstruit.

Évitez d'éditer la présentation pendant une recherche ou un choix de résultat :
PowerPoint ne fournit pas de numéro de révision universel pour chaque modification.
Jarvis refuse les changements détectés, mais ne garantit pas un instantané
atomique face à des modifications simultanées dans PowerPoint ou par un autre
programme. Après une modification, annulez puis relancez la recherche.
Aucun OCR, modèle sémantique ou service cloud n'est utilisé.

Le cache de texte normalisé est limité à 16 777 216 caractères ; au-delà,
la recherche continue sans conserver ce cache pour les requêtes suivantes.
Le contrôle du micro continue en arrière-plan, mais l'interrogation de PowerPoint
est suspendue lorsque le panneau est caché et qu'aucun suivi de répétition ou
point de reprise ne la nécessite. Aucun gain global de CPU n'est garanti.

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

Les fenêtres utilisent les couleurs système, des parcours clavier et des zones
redimensionnables ou défilantes. La mise à l'échelle suit le DPI système au
démarrage ; un déplacement entre des écrans de DPI différents n'offre pas une
adaptation complète par moniteur.

## Vérifier avant de présenter

**Vérifier avant présentation...** contrôle la langue installée, le microphone
sélectionné, l'état de l'écoute, l'activité sonore récente, la présence du
diaporama et l'emplacement de l'indicateur. Le diagnostic n'ouvre pas un second
microphone et ne change pas les slides.

Le silence est un avertissement, pas la preuve d'un microphone défectueux.
La présence d'un périphérique ne prouve pas que la voix est bien reconnue :
le test local reste accessible via **Langue et microphone...**. Le contrôle
d'écran signale les risques, sans garantir que votre partage masque les fenêtres
de Jarvis.

## Parcours courts et longs

Dans **Parcours, budgets et historique...**, configurez plusieurs parcours d'une
même présentation enregistrée, par exemple **5 minutes**, **15 minutes** et
**30 minutes**. Ajoutez les slides voulus, choisissez leur ordre, puis enregistrez.
La durée cible est indicative : le panneau la compare aux budgets des slides,
mais ne choisit pas automatiquement le contenu et ne coupe pas votre intervention.

**Activer** demande confirmation et démarre au premier slide du parcours. Les
commandes suivant/précédent suivent alors cet ordre ; à la fin, Jarvis reste sur
le dernier slide plutôt que de terminer le diaporama. **Parcours complet** rétablit
la navigation habituelle sans déplacement. Le mode questions et les recherches
permettent de sortir temporairement du parcours, puis de reprendre.

Les parcours utilisent les identifiants des slides, pas leurs numéros fragiles.
Ils ne modifient ni les diapositives ni le fichier PowerPoint. Leur activation
est limitée au diaporama courant ; les définitions restent enregistrées localement.

## Raccourcis clavier de secours

Activez **Activer les raccourcis de secours** dans le panneau présentateur.
Ils sont désactivés par défaut et restent utilisables lorsque l'écoute est en
pause ou que le microphone ne fonctionne pas.

| Raccourci | Action |
|---|---|
| Ctrl + Alt + Maj + Droite | Slide suivant |
| Ctrl + Alt + Maj + Gauche | Slide précédent |
| Ctrl + Alt + Maj + Entrée | Reprise |
| Ctrl + Alt + Maj + N | Autre résultat |
| Ctrl + Alt + Maj + B | Écran noir |
| Ctrl + Alt + Maj + S | Réafficher les slides |

Ces combinaisons sont globales tant qu'elles sont activées. Un conflit avec un
autre logiciel est signalé ; Jarvis ne force pas l'enregistrement. Les touches
ordinaires ne sont pas surveillées ou enregistrées. Aucun mode « maintenir pour
parler » n'est ajouté.

## Répétition et données locales

Le chronométrage suit également les changements de slide effectués au clavier
ou à la souris (échantillonnage toutes les 500 ms). Le temps passé en écran noir
est exclu. La fin ou le changement de diaporama arrête la répétition. Une
nouvelle répétition remplace le bilan précédent.
Le temps de veille du PC est également exclu du chronométrage.

Pour une présentation enregistrée, les budgets par slide sont mémorisés et les
bilans de répétition sont conservés localement à l'arrêt, y compris si le
diaporama se termine ou change. Les 20 bilans les plus récents sont conservés.
**Parcours, budgets et historique...** permet de comparer deux répétitions,
slide par slide, et de voir l'écart de durée. L'historique stocke les identifiants,
durées et budgets, pas le texte des slides.

Pour une présentation non enregistrée, le bilan reste uniquement en mémoire :
Jarvis le signale. **Exporter CSV...** permet de l'enregistrer dans un fichier
choisi : titre de présentation, numéro et titre de slide, temps, budget et
dépassement. **Enregistrer le bilan** permet de réessayer après un échec
d'enregistrement local. Aucun audio n'est enregistré.

Les alias sont enregistrés sous `%LOCALAPPDATA%\JarvisPowerPoint`, sans modifier
le fichier PowerPoint. Leur association dépend de l'emplacement du fichier :
après un déplacement ou un renommage de la présentation, recréez ses alias.
La langue et le microphone sont mémorisés dans
`HKEY_CURRENT_USER\Software\JarvisPowerPoint`.

La reconnaissance vocale reste entièrement locale. Aucune donnée audio n'est
envoyée vers un service en ligne.

## Diagnostic local exportable

**Diagnostic local...** affiche un aperçu et propose un export texte volontaire.
Il contient les versions techniques, quelques états (langue, écoute, diaporama,
écrans, raccourcis) et au maximum 100 événements : catégorie, date, succès/échec,
type d'erreur et code technique.

Il ne contient ni audio, ni phrases reconnues, ni titres/contenus de slides, ni
noms d'alias, noms d'appareils ou chemins de fichiers. Les événements ne sont pas
enregistrés automatiquement sur disque et aucun rapport n'est envoyé.

## Mises à jour facultatives

**Mises à jour... → Rechercher sur GitHub** consulte uniquement les versions
publiques du dépôt `sebplace/jarvis-powerpoint`. Il n'y a aucune recherche réseau
automatique. Les notes de version sont affichées comme texte, pas exécutées.

Après votre confirmation, Jarvis télécharge la nouvelle version, contrôle sa
taille, son SHA-256 fourni par GitHub et sa version intégrée. L'installation est
refusée pendant un diaporama ou une répétition. Jarvis se ferme puis redémarre ;
les réglages, alias et profils restent à leur emplacement.

Une copie de l'ancien exécutable est conservée dans le même dossier. Si le
remplacement ou le lancement échoue, l'assistant tente une restauration et
indique comment récupérer l'application. Aucun droit administrateur n'est demandé.
La vérification SHA-256 ne constitue pas une signature Authenticode.

Dans une installation MSI, ce remplacement portable est désactivé, y compris
dans l'assistant technique de mise à jour. Le dialogue renvoie vers le service
informatique sans contacter GitHub. L'organisation conserve la gestion des
versions, réparations et désinstallations avec Windows Installer.

Les fichiers OneDrive standard sont pris en charge. Les liens symboliques,
jonctions, chemins UNC et points de réanalyse inconnus sont refusés.
L'assistant conserve ses résultats techniques sous
`%LOCALAPPDATA%\JarvisPowerPoint\Updates`. Ces résultats peuvent contenir des
chemins locaux ; ils ne font pas partie du diagnostic expurgé décrit plus haut.

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

## MSI et déploiement professionnel

Le package `JarvisPowerPoint-1.5.0-x86.msi` installe l'application dans
`[ProgramFilesFolder]\Jarvis PowerPoint` et crée un raccourci Démarrer pour tous
les utilisateurs. Il ne lance pas l'application sous SYSTEM, ne crée aucun
service ou démarrage automatique et conserve les réglages, alias et profils
personnels lors de la désinstallation.

Le MSI est de type x86, mais l'exécutable reste **AnyCPU / .NET Framework**.
Prévoyez .NET Framework 4.8+, PowerPoint et ses composants d'interopérabilité,
les langues vocales Windows et l'accès au microphone. Les configurations
ARM64/Office émulé demandent une qualification spécifique avant déploiement.

```powershell
.\build-msi.ps1 -Executable .\bin\candidate\JarvisPowerPoint.exe
msiexec.exe /i "JarvisPowerPoint-1.5.0-x86.msi" /qn /norestart
msiexec.exe /fa {B39CDD36-C72B-41A6-9CE2-E3D59B53EB21} /qn /norestart
msiexec.exe /x {B39CDD36-C72B-41A6-9CE2-E3D59B53EB21} /qn /norestart
```

Le build utilise WiX 3.14.1 épinglé ; si cet outil manque, le script indique
comment restaurer sa copie locale. Pour Intune, utilisez le contexte **System**
et la détection MSI par ProductCode ci-dessus, avec une règle de version adaptée.
Le code `0` indique le succès, `3010` demande un redémarrage et `1618` un nouvel
essai. Fermez Jarvis avant les opérations de maintenance ; aucun processus n'est
tué pour forcer un remplacement. Une copie portable existante n'est pas migrée.

Les instructions complètes (journaux, codes retour, registre, mises à niveau,
qualification et signature) sont dans [installer/DEPLOYMENT.txt](installer/DEPLOYMENT.txt).
Une version de production doit être qualifiée par l'IT : inspection du package
et extraction administrative ne remplacent pas une installation réelle en VM.
Lorsque le poste bloque la validation ICE de WiX, le build peut produire un
artefact de préparation explicitement marqué avec `-SkipIceValidation` ;
la validation ICE reste requise sur un poste de build autorisé avant diffusion.

Le script `sign-release.ps1` prépare la signature Authenticode avec **un
certificat déjà délivré** et un `signtool` existant. Signez d'abord l'exécutable,
construisez le MSI avec cet exécutable signé, puis signez le MSI. Sans ces moyens,
les artefacts restent non signés. Une signature ne garantit ni l'approbation de
l'organisation ni l'absence d'avertissement SmartScreen.

## Licence

Ce projet est distribué sous [licence MIT](LICENSE).
