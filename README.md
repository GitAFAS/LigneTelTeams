# LigneTelTeams
Gestion des lignes téléphoniques Microsoft Teams avec un excel


						
![image](https://github.com/user-attachments/assets/64f08472-acae-414a-84c3-6fb939f0bfdb)


Utilisation : 

<b>Bouton « PUSH TO PS »</b>

Il lance les 5 commandes powershell suivantes : 
-		Connect-MicrosoftTeams : ouverture de la console PSTeams avec droit admin demandé
		Set-CsPhoneNumberAssignment  : Pour ajout du numéro + direct routing + location
		Grant-CsOnlineVoiceRoutingPolicy : Pour ajout de la Policy
		Grant-CsTenantDialPlan : Pour ajout du dialPlan
		Grant-CsCallingLineIdentity : Pour ajout d’une calling Line Identity 

Il a donc besoin de : Mail, SDA BRUTE, CallingLineIdentity (pas obligatoire), Site ID, policy et Dialplan


<b>Bouton « GET FROM PS » </b>

Il va chercher sur Teams les informations concernant un utilisateur. Le seul champ obligatoire est le Mail. Il récupère les informations suivantes, demande confirmation si changement, et remplis le fichier tout seul : 

-		PhoneNumber qui va dans SDA BRUTE
		OnlineVoiceRoutingPolicy qui va dans Policy
		TenantDialPlan qui va dans DialPlan
		CallingLineIdentity qui va dans CallingLineIdentity
