# HappyEnd - Parce qu'il y a toujours une bonne fin

Ce programme permet d'extraire tous les emails et autres éléments (rendez-vous, contacts, tâches) de Microsoft Outlook et de les sauvegarder sur votre ordinateur dans différents formats universels.

De plus, il permet de contourner les sécurités mises en place par les entreprises pour l'encryption des emails et offre également différents formats de sortie, ce qui peut être utile en cas de perte d'accès à Exchange ou si vos droits ont été révoqués. (par pur hasard)
C'qui est proposé ailleurs ne me satisfait pas.

## Fonctionnalités

- Extraction de tous les emails de tous les dossiers Outlook
- Sauvegarde des emails dans plusieurs formats
- Sauvegarde des contacts au format `.vcf` (vCard) compatible avec la plupart des gestionnaires de contacts
- Extraction des pièces jointes
- Sauvegarde des rendez-vous, contacts et tâches
- Prise en charge des fichiers PST et autres magasins de données

## Prérequis

- Microsoft Outlook doit être installé sur votre ordinateur
- .NET 9.0 ou supérieur
- Accès à votre compte Outlook

## Formats des fichiers exportés

Le programme exporte les éléments Outlook dans les formats suivants :

### Emails
- `.msg` : Format natif d'Outlook
- `.eml` : Format standard d'email
- `.txt` : Fichier texte brut avec métadonnées (De, À, Cc, Objet, Date)
- `.html` : Version HTML de l'email

### Contacts
- `.msg` : Format natif d'Outlook
- `.vcf` : Format vCard standard
- `.txt` : Fichier texte avec les informations du contact

### Rendez-vous et Tâches
- `.msg` : Format natif d'Outlook
- `.txt` : Fichier texte avec toutes les informations importantes

## Dossiers extraits

Le programme extrait les emails des dossiers suivants :
- Boîte de réception (Inbox)
- Éléments envoyés (Sent Items)
- Brouillons (Drafts)
- Éléments supprimés (Deleted Items)
- Boîte d'envoi (Outbox)
- Courrier indésirable (Junk Email)
- Tous les sous-dossiers de ces dossiers principaux
- Tous les dossiers des fichiers PST et autres magasins de données


- Le programme ne peut pas extraire les emails des comptes Exchange Online si l'authentification moderne est activée sans qu'Outlook soit configuré
