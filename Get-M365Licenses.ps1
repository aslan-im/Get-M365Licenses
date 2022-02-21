#Requires -module MSOnline
#Requires -module ImportExcel
#Requires -Version 5
Import-Module MSOnline
Import-Module ImportExcel

try {
    Connect-MsolService -ErrorAction Stop
}
catch {
    Write-Error "There is no active connection to MSOL. $($_.Exception)"
    break
}


Write-Output "Collecting all available licenses"
try {
    $AllLicenses = Get-MsolAccountSku 
}
catch {
    Write-Error "Can't collect licenses. Check your permissions"
}

#List of SkuIDs and friendly names
$LicensesDictionary = @{
    'STREAM'                                = 'Microsoft Stream Trial';
    'POWER_BI_PRO'                          = 'Power BI Pro';
    'SPZA_IW'                               = 'App Connect';
    'WINDOWS_STORE'                         = 'Windows Store for Business';
    'FLOW_FREE'                             = 'Microsoft Flow Free';
    'MICROSOFT_BUSINESS_CENTER'	            = 'Microsoft Business Center';
    'MCOEV'                                 = 'Microsoft 365 Phone System';
    'CCIBOTS_PRIVPREV_VIRAL'                = 'Dynamics Bots Trial';
    'FORMS_PRO'                             = 'Forms Pro Trial';
    'POWERAPPS_VIRAL'                       = 'Microsoft PowerApps Plan 2 Trial';
    'MCOCAP'                                = 'Common Area Phone';
    'CRMTESTINSTANCE'                       = 'Microsoft Dynamics CRM Test Instance';
    'DYN365_ENTERPRISE_PLAN1'               = 'Dynamics 365 Customer Engagement Plan Enterprise Edition';
    'MEETING_ROOM'                          = 'Microsoft Teams Rooms Standard';
    'POWER_BI_STANDARD'                     = 'Power BI (free)';
    'MCOPSTNC'                              = 'Communications Credits';
    'ADALLOM_STANDALONE'                    = 'Microsoft Cloud App Security';
    'TEAMS_EXPLORATORY'                     = 'Microsoft Teams Exploratory';
    'MCOMEETADV'                            = 'Microsoft 365 Audio Conferencing';
    'CRMINSTANCE'                           = 'Microsoft Dynamics CRM Instance';
    'SPE_E3'                                = 'Microsoft 365 E3';
    'MCOPSTN2'                              = 'Microsoft 365 Domestic and International Callin plan';
    'CRMSTORAGE'                            = 'Microsoft Dynamics CRM Storage';
    'RIGHTSMANAGEMENT_ADHOC'                = 'Rights Management Adhoc';
    'STANDARDPACK'                          = 'Office 365 E1';
    'EMSPREMIUM'                            = 'Enterprise Mobility + Security E5';
    'FLOW_P1'                               = 'Microsoft Flow Plan 1';
    'MICROSOFT_REMOTE_ASSIST'	            = 'Dynamics 365 Remote Assist';
    'ENTERPRISEPREMIUM'	                    = 'Office 365 E5';
    'FLOW_PER_USER'	                        = 'Power Automate per user plan';
    'DYN365_AI_SERVICE_INSIGHTS'	        = 'Trial Dynamics 365 Customer Service Insights';
    'POWERFLOW_P1'	                        = 'Microsoft Power Automate Plan 1';
    'ENTERPRISEPACK'	                    = 'Office 365 E3';
    'M365_E5_SUITE_COMPONENTS'	            = 'Microsoft 365 E5 Suite features';
    'PROJECTESSENTIALS'	                    = 'Project Online Essentials';
    'M365_F1'	                            = 'Microsoft 365 F1';
    'DESKLESSPACK'	                        = 'Office 365 F1';
    'OFFICE365_MULTIGEO'	                = 'Multi-Geo Capabilities in Office 365';
    'PROJECT_P1'	                        = 'Project Plan 1';
    'PROJECTPREMIUM'	                    = 'Project Plan 5';
    'PBI_PREMIUM_P1_ADDON'	                = 'Power BI Premium Plan 1 Addon';
    'EXCHANGESTANDARD'	                    = 'Exchange Online Plan 1';
    'DYN365_ENTERPRISE_P1_IW'	            = 'Dynamics 365 P1 Trial for Information Workers';
    'DYN365_ENTERPRISE_CUSTOMER_SERVICE'    = 'Dynamics 365 for Customer Service Enterprise Edition';
    'WIN_DEF_ATP'	                        = 'Microsoft Defender Advanced Threat Protection';
    'POWERFLOW_P2'	                        = 'Microsoft Power Apps Plan 2';
    'POWERAPPS_PER_USER'	                = 'Power Apps Per User Plan';
    'EMS'	                                = 'Enterprise Mobility + Security E3';
    'PBI_PREMIUM_P2_ADDON'	                = 'Power BI Premium Plan 2 Addon';
    'M365_F1_COMM'	                        = 'Microsoft 365 F1';
    'AAD_PREMIUM'	                        = 'Azure Active Directory Premium P1';
    'PROJECTPROFESSIONAL'	                = 'Project Online Professional';
    'EXCHANGEENTERPRISE'	                = 'Exchange Online Plan 2';
    'SPE_F1'	                            = 'Microsoft 365 F3';
    'WORKPLACE_ANALYTICS'	                = 'Microsoft Workplace Analytics';
    'POWERAPPS_PER_APP'	                    = 'Power Apps Per Application';
    'DYN365_ENTERPRISE_TEAM_MEMBERS'	    = 'Dynamics 365 for Team Members Enterprise Edition';
    'POWERAPPS_PER_APP_IW'                  = 'PowerApps per app baseline access';
    'PHONESYSTEM_VIRTUALUSER'               = 'Microsoft 365 Phone System - Virtual Usuer';
    'EXCHANGEARCHIVE_ADDON'                 = 'Exchange Online Archiving addon';
    'RIGHTSMANAGEMENT'                      = 'Azure Information Protection Premium P1';
    'CDSAICAPACITY'                         = 'AI Builder Capacity add-on';
    'STREAM_STORAGE'                        = 'Microsoft Stream Storage Add-On (500 GB)';
    'VISIOCLIENT'                           = 'Visio Online Plan 2';
    'M365_INFO_PROTECTION_GOVERNANCE'       = 'Microsoft 365 E5 Information Protection and Governance';
    'THREAT_INTELLIGENCE'                   = 'Microsoft Defender for Office 365 (Plan 2)';
    'FLOW_BUSINESS_PROCESS'                 = 'Power Automate per flow plan'
    'DYN365_TEAM_MEMBERS'                   = 'Dynamics 365 Team Members'
    'REMOTE_ASSIST_DEVICE'                  = 'Dynamics 365 Remote Assist Device';
    'MCOPSTN1'                              = 'Skype for Business PSTN Domestic Calling';
    'AAD_PREMIUM_P2'                        = 'Azure Active Directory Premium P2';
    'MTR_PREM'                              = 'Teams Rooms Premium';
    'OFFICESUBSCRIPTION'                    = 'Microsoft 365 Apps for Enterprise';
    'SPE_E5'                                = 'Microsoft 365 E5';
    'PROJECT_PLAN1_DEPT'                    = 'Project Plan 1 (for Department)';
    'ADV_COMMS'                             = 'Advanced Communications';
    'POWERAUTOMATE_ATTENDED_RPA'            = 'Power Automate per user with attended RPA plan';
    'CDS_DB_CAPACITY'                       = 'Common Data Service Database Capacity';
    'DYN365_ENTERPRISE_SALES'               = 'Dynamics 365 For Sales and Customer Service Enterprise Edition';
    'STREAM_P2'                             = 'Microsoft Stream Plan 2';
    'CDS_LOG_CAPACITY'                      = 'Common Data Service Log Capacity';
    'ATP_ENTERPRISE'                        = 'Microsoft Defender for Office 365 (Plan 1)';
    'CDS_API_CAPACITY'                      = 'Common Data Service API Capacity';
    'POWERAPPS_DEV'                         = 'Power Apps for Developer';
}

#Getting first SKU for generating tenant name
$firstSKU = ($AllLicenses[0]).AccountSkuId
$OrganizationName,$sku = $firstSKU -split ':'
[string]$CurrentDate = get-date -Format 'yyyy.MM.dd'
$ReportName = $CurrentDate + '_' + $OrganizationName + '.xlsx'

#Formatting the result and exporting to .xlsx
$LicensesSelection = @(
    @{
        L = 'License (Friendly name)'; 
        E = {
            
            $([string]$CompanyName, [string]$SkuID = $_.AccountSkuId -split ':'
            "$($LicensesDictionary[$SkuID])")
        }
    }, 
    @{
        L = 'SKU Name';
        E = {
            $([string]$CompanyName, [string]$SkuID = $_.AccountSkuId -split ':'
            $SkuID)
        }
    },
    @{
        L = 'Valid Licenses'; 
        E = {
            $($($_.ActiveUnits) - $($_.ConsumedUnits))
        }
    },
    @{
        L='Assigned License'; 
        E={$($_.ConsumedUnits)}},
    @{
        L = 'Licenses Total'; 
        E = {$_.ActiveUnits}
    }
)

$LicensesToExport = $AllLicenses | Select-Object $LicensesSelection | Sort-Object 'SKU Name'

try {
    $LicensesToExport | Export-Excel -Path $ReportName -AutoSize -AutoFilter -TableStyle Medium2
    Write-Verbose "Report generated and exported to $ReportName"
}
catch {
    Write-Error "Can't export report. $($_.Exception)" 
}


