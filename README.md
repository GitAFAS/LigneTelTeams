# LigneTelTeams
Gestion des lignes téléphoniques Microsoft Teams avec un excel


						
![image](https://github.com/user-attachments/assets/64f08472-acae-414a-84c3-6fb939f0bfdb)


### Description des colonnes du document

Les colonnes en $${\color{red}rouge}$$ sont obligatoire, celles en $${\color{green}verte}$$ sont optionnelles, celles en bleu $${\color{lightblue}bleu}$$ ne sont pas utilisés par les scripts.

| Nom de la colonne  | Description |
| ------------- | ------------- |
| Site  | Site géographique de l'utilisateur  |
| Nom  | Nom de l'utilisateur  |
| Prénom | Prénom de l'utilisateur
| Licence | Licence associée à l'utilisateur |
| Mail | Il s'agit du _UserPrincipalName_ utilisé avec le paramètre **-Identity** | 
| SDA | Numéro de téléphone en format human readable | 
| SDA BRUTE | Numéro de téléphone en format accepté par teams utilisé avec le paramètre -PhoneNumber | 
| CallingLineIdentity | [Stratégie d'identification de l'appelant](https://admin.teams.microsoft.com/policies/callinglineid) (Abrégé en "CLI") comme elle apparait dans la liste colonne **Nom** | 
| Site ID | [Adresse d'Urgence](https://admin.teams.microsoft.com/locations) de l'emplacement au format xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx | 
| Policy | [Stratégie de routage des communicatons vocales](https://admin.teams.microsoft.com/policies/teamsonlinevoicerouting) comme elle apparait dans la liste colonne **Nom**  | 
| DialPlan | [Plan de numérotation](https://admin.teams.microsoft.com/policies/teamsdialplan) comme il apparait dans la liste colonne **Nom**| 
| Get From PS | Horodatage de l'execution de la macro Get_From_PS sur cette ligne | 
| Push To PS | Horodatage de l'execution de la macro Push_To_PS sur cette ligne | 
| Options | Champ libre | 


# Utilisation des macros : 

## Bouton « PUSH TO PS »

Il lance les 5 commandes powershell suivantes : 
-		Connect-MicrosoftTeams : ouverture de la console PSTeams avec droit admin demandé
		Set-CsPhoneNumberAssignment  : Pour ajout du numéro + direct routing + location
		Grant-CsOnlineVoiceRoutingPolicy : Pour ajout de la Policy
		Grant-CsTenantDialPlan : Pour ajout du dialPlan
		Grant-CsCallingLineIdentity : Pour ajout d’une calling Line Identity 

Il a donc besoin de : **Mail**, **SDA BRUTE**, **CallingLineIdentity** (pas obligatoire), **Site ID**, **policy** et **Dialplan**


## Bouton « GET FROM PS » 

Il va chercher sur Teams les informations concernant un utilisateur. Le seul champ obligatoire est le **Mail**. Il récupère les informations suivantes, demande confirmation si changement, et remplis le fichier tout seul : 

-		PhoneNumber qui va dans SDA BRUTE
		OnlineVoiceRoutingPolicy qui va dans Policy
		TenantDialPlan qui va dans DialPlan
		CallingLineIdentity qui va dans CallingLineIdentity
