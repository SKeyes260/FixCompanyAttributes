# PURPOSE
# Identify all systems in AD that have a value for companyAttributeMachineType set to "NON-STANDARD" 
# modify any records matched SCCM NetBoisName and/or SMSUniqueID whose data does not match AD
# Create a DDR record for SCCM to update its discovery data with the values from AD

param([string]$SiteServer, [string]$SiteCode, [string]$InstanceName )

PowerShell { Import-Module -Global ActiveDirectory -Force }

# Fill the objSMSSites array from the CAS WMI
Function Get-SMSSites() {
[CmdletBinding()]   
PARAM 
(   [Parameter(Position=1)] $SiteServer,
    [Parameter(Position=2)] $SiteCode )   
    
    $objSMSSites = Get-WmiObject -ComputerName $SiteServer -Namespace ("root\sms\Site_"+$SiteCode) -Query "SELECT * FROM SMS_SITE Order By SiteCode"
    Return $objSMSSites
}


# Fill the objDomains array with SQL table data from SCCM_EXT 
Function Get-ADDomains() {
[CmdletBinding()]   
PARAM 
(   [Parameter(Position=1)] $SQLServer,
    [Parameter(Position=2)] $ExtDiscDBName,
    [Parameter(Position=3)] $DomainsTableName
)   

    $objADDomains = @()
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server=$SQLServer;Database=$ExtDiscDBName;Integrated Security=True"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = "select * from $DomainsTableName"
    $SqlCmd.Connection = $SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $objADDomains = New-Object System.Data.Datatable
    $NumRows = $SqlAdapter.Fill($objADDomains)
    $SqlConnection.Close()
    Return $objADDomains
}

Function Create-DDR() {
[CmdletBinding()]   
PARAM 
(   [Parameter(Position=1)] $SiteCode,
    [Parameter(Position=2)] $ResourceID,
    [Parameter(Position=3)] $NetBiosName,   
    [Parameter(Position=4)] $SMSUniqueIdentifier,
    [Parameter(Position=5)] $DistinguishedName,
    [Parameter(Position=6)] $Category,
    [Parameter(Position=7)] $Type  )  
        
    $SMSDisc.DDRNew("System","SA_EXT_Disc",$SiteCode) 

    If (($SMSUniqueIdentifier) -AND ($SMSUniqueIdentifier -NE '')  -AND ($SMSUniqueIdentifier -NE ' '))  {
        If ($SMSUniqueIdentifier.length -ge 36) { 
            $SMSUID = $SMSUniqueIdentifier.substring($SMSUniqueIdentifier.length - 36, 36)
            $SMSDisc.DDRAddString("SMS Unique Identifier", $SMSUniqueIdentifier, 64,  $ADDPROP_GUID + $ADDPROP_KEY)
         } 
         ELSE { 
             $SMSUID = "NOGUID"  
         }
    }

    If ($NetBiosName       -AND $NetBiosName         -NE $Null) { $SMSDisc.DDRAddString("Netbios Name", $NetBiosName, 16,  $ADDPROP_NAME + $ADDPROP_KEY) } 
    If ($DistinguishedName -AND $DistinguishedName   -NE $Null) { $SMSDisc.DDRAddString("Distinguished Name", $DistinguishedName,  256, $ADDPROP_NONE)  }
    If ( $Category -NE $Null) { $SMSDisc.DDRAddString("companyAttributeMachineCategory",  $Category, 32,  $ADDPROP_NONE)  }
    If ( $Type     -NE $Null) { $SMSDisc.DDRAddString("companyAttributeMachineType",      $Type, 32,  $ADDPROP_NONE)  }

    $Result = $SMSDisc.DDRWrite($DDRTempFolder+$SiteCode+"-"+$NetBiosName+"-"+$SMSUID+".DDR")
    $TestFile = Get-Item -LiteralPath  ($DDRTempFolder+$SiteCode+"-"+$NetBiosName+"-"+$SMSUID+".DDR")
    Return $TestFile
}


Function Log-Append () {
[CmdletBinding()]   
PARAM 
(   [Parameter(Position=1)] $strLogFileName,
    [Parameter(Position=2)] $strLogText )
    
    $strLogText = ($(get-date).tostring()+" ; "+$strLogText.ToString()) 
    Out-File -InputObject $strLogText -FilePath $strLogFileName -Append -NoClobber
}


Function Log-SQLDDRChange() {
PARAM 
(   [Parameter(Position=1)] $SQLServer,
    [Parameter(Position=2)] $ExtDiscDBName, 
    [Parameter(Position=3)] $LoggingTableName, 
    [Parameter(Position=4)] $ComputerName,
    [Parameter(Position=5)] $AD_DN,
    [Parameter(Position=6)] $AD_Category,
    [Parameter(Position=7)] $AD_Type,
    [Parameter(Position=12)] $SCCM_DN,
    [Parameter(Position=13)] $SCCM_Category,
    [Parameter(Position=14)] $SCCM_Type,
    [Parameter(Position=15)] $SCCM_SiteCode,
    [Parameter(Position=16)] $SMSClientGUID)

    If (!$SQLConnection) {
        $SQLConnection =  New-Object System.Data.SqlClient.SqlConnection  
        $SQLConnection.ConnectionString = "Server=$SQLServer;Database=$ExtDiscDBName;Integrated Security=True"
        $SQLConnection.Open()
    }
    $cmd = $SQLConnection.CreateCommand()
    $cmd.CommandText ="INSERT INTO $LoggingTableName  (ComputerName,AD_DN,AD_Category,AD_Type,SCCM_DN,SCCM_Category,SCCM_Type,SCCM_SiteCode,SMSClientGUID,DiscoveryMethod) 
                        VALUES( '$ComputerName','$AD_DN','$AD_Category','$AD_Type','$SCCM_DN','$SCCM_Category','$SCCM_Type','$SCCM_SiteCode','$SMSClientGuid','$InstanceName' );"
    $Result = $cmd.ExecuteNonQuery()
    $Result = $SQLConnection.Close
}


Function Get-SCCMComputer() {
[CmdletBinding()]   
PARAM 
(   [Parameter(Position=1)] $SQLServer,
    [Parameter(Position=2)] $SCCMDBName,   
    [Parameter(Position=3)] $ComputerName,
    [Parameter(Position=4)] $DistinguishedName )    

        $objSCCMClients = @()
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = "Server=$SQLServer;Database=$SCCMDBName;Integrated Security=True"
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText =  "SELECT     TOP (100) PERCENT dbo.v_RA_System_SMSAssignedSites.SMS_Assigned_Sites0, dbo.v_R_System.ResourceID, dbo.v_R_System.companyAttributeMachineCa0, 
                                dbo.v_R_System.companyAttributeMachineTy0, dbo.v_R_System.Distinguished_Name0, dbo.v_R_System.Full_Domain_Name0, dbo.v_R_System.Netbios_Name0, 
                                dbo.v_R_System.Resource_Domain_OR_Workgr0, dbo.v_R_System.SMS_Unique_Identifier0
                                FROM  dbo.v_R_System 
                                LEFT OUTER JOIN dbo.v_RA_System_SMSAssignedSites ON dbo.v_R_System.ResourceID = dbo.v_RA_System_SMSAssignedSites.ResourceID
                                WHERE     (dbo.v_R_System.Netbios_Name0 = '$ComputerName')"
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
        $objSCCMClients = New-Object System.Data.DataTable
        $ReturnedRows = $SqlAdapter.Fill($objSCCMClients)
        $SqlConnection.Close()
        Return $objSCCMClients
}



##############################
#  MAIN
##############################

#Define standard static variables
$ADDPROP_NONE  = 0x0
$ADDPROP_GUID  = 0x2
$ADDPROP_KEY   = 0x8
$ADDPROP_ARRAY = 0x10
$ADDPROP_NAME  = 0x44


#Define environment specific  variables
$LoggingFolder      = "C:\TEMP\FixCompanyAttributes\"
$DDRTempFolder = ($LoggingFolder+"TempDDRs\")
$ExtDiscDBName = "SCCM_EXT"
$DomainsTableName = "tbl_ExtDiscDomains"
$LoggingTableName   = "tbl_ExtDiscLogging"
$DaysInactive = 90  
$TodaysDate = Get-Date
$InactivityDate = (Get-Date).Adddays(-($DaysInactive))

# Set parameter defaults
$PassedParams =  (" -SiteServer "+$SiteServer+" -SiteCode "+$SiteCode+" -SQLServer "+$SQLServer+" -InstanceName "+$InstanceName)
If (!$SiteServer)   { $SiteServer   = "XSNW10S629K"                   }
If (!$SiteCode)     { $SiteCode     = "F01"                           }
If (!$InstanceName) { $InstanceName = "FixCompanyAttributes"    }

# Lookup the SCCM site definition to populate global variables
$objSiteDefinition = Get-WmiObject -ComputerName $SiteServer -Namespace ("root\sms\Site_"+$SiteCode) -Query ("SELECT * FROM sms_sci_sitedefinition WHERE SiteCode = '"+$SiteCode+"'")
$SCCMDBName = $objSiteDefinition.SQLDatabaseName
$SQLServer = $objSiteDefinition.SQLServerName

# Get the logging prepared
If (!(Test-Path $LoggingFolder))  { $Result = New-Item $LoggingFolder -type directory }
If (!(Test-Path $DDRTempFolder))  { $Result = New-Item $DDRTempFolder -type directory }
$LogFileName = ($LoggingFolder+$InstanceName+"-"+$TodaysDate.Year+$TodaysDate.Month.ToString().PadLeft(2,"0")+$TodaysDate.Day.ToString().PadLeft(2,"0")+".log")
    
# Set variables that use parameters
$DefaultDDRSiteCode = $SiteCode # used for clients with no assigned site code
$DDRTargetFolder = ("\sms_"+$SiteCode+"\inboxes\auth\ddm.box")   #must include trailing \

# Define global Arrays
$objSMSSites = @()
$objADDomains = @()
$objLogging = @()
$ADComputers = @()    
$SCCMComputers = @()

Log-Append -strLogFileName $LogFileName -strLogText ("Script Started using command line parameters : "+$PassedParams )
Log-Append -strLogFileName $LogFileName -strLogText ("The actual values being used are : -SiteServer "+$SiteServer+" -SiteCode "+$SiteCode+" -SQLServer "+$SQLServer+" -InstanceName "+$InstanceName)

#Load the SCCM SDK DLL
Log-Append -strLogFileName $LogFileName -strLogText ("Creating an instance of the com object SMSResGen.SMSResGen.1" )
If (!$SMSDisc) {
    Try { &regsvr32 /s ($DDRTempFolder+"\OldDLLs\smsrsgen.dll")
        $SMSDisc = New-Object -ComObject "SMSResGen.SMSResGen.1" }
    Catch { 
        Try {
            &regsvr32 /s ($DDRTempFolder+"\NewDLLs\smsrsgen.dll")
            $SMSDisc = New-Object -ComObject "SMSResGen.SMSResGen.1" 
        }
        Catch { Log-Append -strLogFileName $LogFileName -strLogText "Failed to load COM object SMSResGen.SMSResGen.1" }
    }
}

#Get a list of the  site codes in SCCM 
Log-Append -strLogFileName $LogFileName -strLogText ("Getting SCCM primary site information")
$objSMSSites = Get-SMSSItes -SiteServer $SiteServer  -SiteCode $SiteCode
ForEach ($Site in $objSMSSites){ Log-Append -strLogFileName $LogFileName -strLogText (" - "+$Site.ServerName+"    "+$Site.SiteCode+"    "+$Site.InstallDir) }

# Get a list of all of the discoverable domains from SQL Ext table
Log-Append -strLogFileName $LogFileName -strLogText ("Getting list of SCCM discovered domains from SQL")
$objADDomains = Get-ADDomains -SQLServer $SQLServer -ExtDiscDBName $ExtDiscDBName -DomainsTableName $DomainsTableName

# Create a list of all computer objects in all discovered domains that have company attributes set in AD 
ForEach ($objDomain in $objADDomains ) {
    #Log-Append -strLogFileName $LogFileName -strLogText ("Searching the domain named "+$objDomain.DomainNameFQDN+" for computers with values set for the company attributes using DC named "+$objDomain.PDCFQDN)
    TRY {  
        $ADComputers = Get-ADComputer -Server $objDomain.PDCFQDN -Filter  "CompanyAttributeMachineCategory -eq 'NON-STANDARD' -OR CompanyAttributeMachineCategory -eq 'CLEARED'"    -Properties name,distinguishedName,companyAttributeMachineCategory,companyAttributeMachineType
        Log-Append -strLogFileName $LogFileName -strLogText ("Identified "+$ADComputers.Count+" computers in domain "+$objDomain.DomainNameFQDN+" with company attribute information.")
    }
    CATCH { Log-Append -strLogFileName $LogFileName -strLogText ("Failed to connect to domain named "+$objDomain.DomainNameFQDN+" using DC named "+$objDomain.PDCFQDN)}

    # Iterate through the list of computers from AD and see if SCCM has a matching record
    # If the company attributes for the SCCM record does not match AD then write an SCCM DDR with the AD information
    ForEach ($ADComputer in $ADComputers) { 
        Log-Append -strLogFileName $LogFileName -strLogText  ("Checking SCCM for a computer named "+$ADComputer.Name)
        $FoundMatch = $False 
        $SCCMComputers = Get-SCCMComputer  -SQLServer $SQLServer -SCCMDBName $SCCMDBName -ComputerName $ADComputer.Name -DistinguishedName $ADComputer.DistinguishedName
        Foreach ( $SCCMComputer in $SCCMComputers) {  
            $FoundMatch = $True
            $NeedsUpdate = $False
            
            # If Distinguished name is missing then create a DDR
            If ($SCCMComputer.Distinguished_Name0 -eq $Null -OR $SCCMComputer.Distinguished_Name0 -eq ''  )  { 
                Log-Append -strLogFileName $LogFileName -strLogText  ("SCCM Computer "+$SCCMComputer.Netbios_Name0+" does not have a Distinguished_Name0 value") 
                $NeedsUpdate = $True 
            }

            # Check to see if the CompanyAttributeMachineCategory value does not match
            If ( $SCCMComputer.CompanyAttributeMachineCa0 -ne $ADComputer.CompanyAttributeMachineCategory ) { 
                Log-Append -strLogFileName $LogFileName -strLogText ("SCCM Machine Category("+$SCCMComputer.CompanyAttributeMachineCa0+") does not match AD value("+$ADComputer.companyAttributeMachineCategory+")")
                If ( ($SCCMComputer.CompanyAttributeMachineTy0 -ne $ADComputer.CompanyAttributeMachineType) -AND $ADComputer.CompanyAttributeMachineType -eq $Null ) {
                    $ADComputer.CompanyAttributeMachineType = ''
                }
                $NeedsUpdate = $True 
            }
                 
            # SCCM data needs updating                            
            If ( $NeedsUpdate -eq $True ) { 
                If ($SCCMComputer.SMS_Assigned_Sites0.ToString().length -eq 3 ) {
                    $UseSiteCode =  $SCCMComputer.SMS_Assigned_Sites0
                } ELSE { $UseSiteCode = $SiteCode }
                Log-Append -strLogFileName $LogFileName -strLogText ("Found "+$SCCMComputer.NetBios_Name0+" in SCCM whose data does not match AD")
                Log-Append -strLogFileName $LogFileName -strLogText ("Creating DDR:"+$DDRTempFolder+$UseSiteCode+"-"+$SCCMComputer.NetBios_Name0+"-"+$SCCMComputer.SMS_Unique_Identifier0+".ddr")
                Log-Append -strLogFileName $LogFileName -strLogText ("- SiteCode: "+$UseSiteCode)
                Log-Append -strLogFileName $LogFileName -strLogText ("- DDR NetBiosName: "+$SCCMComputer.NetBios_Name0)
                Log-Append -strLogFileName $LogFileName -strLogText ("- DDR SMSUniqueIdentifier: "+$SCCMComputer.SMS_Unique_Identifier0)
                Log-Append -strLogFileName $LogFileName -strLogText ("- DDR DistinguishedName: "+$ADComputer.DistinguishedName)
                Log-Append -strLogFileName $LogFileName -strLogText ("- DDR Category: "+$ADComputer.CompanyAttributeMachineCategory)
                Log-Append -strLogFileName $LogFileName -strLogText ("- DDR Type: "+$ADComputer.CompanyAttributeMachineType)
                Log-Append -strLogFileName $LogFileName -strLogText ("- Old SCCM Category: "+$SCCMComputer.CompanyAttributeMachineCa0)
                Log-Append -strLogFileName $LogFileName -strLogText ("- Old SCCM Type: "+$SCCMComputer.CompanyAttributeMachineTy0)

                $DDRFile = Create-DDR  -SiteCode $UseSiteCode -ResourceID  $SCCMComputer.ResourceID -NetBiosName $SCCMComputer.NetBios_Name0 -SMSUniqueIdentifier $SCCMComputer.SMS_Unique_Identifier0  -DistinguishedName  $ADComputer.DistinguishedName -Category  $ADComputer.CompanyAttributeMachineCategory  -Type $ADComputer.CompanyAttributeMachineType 
                $LogResult = Log-SQLDDRChange -SQLServer $SQLServer -ExtDiscDBName $ExtDiscDBName -LoggingTableName $LoggingTableName -ComputerName $ADComputer.name -AD_DN $ADComputer.DistinguishedName -AD_Category $ADComputer.companyAttributeMachineCategory -AD_Type $ADComputer.companyAttributeMachineType -SCCM_DN $SCCMComputer.Distinguished_Name0 -SCCM_Category $SCCMComputer.companyAttributeMachineCa0 -SCCM_Type $SCCMComputer.companyAttributeMachineTy0 -SMSClientGUID $SCCMComputer.SMS_Unique_Identifier0 -SCCM_SiteCode $SCCMComputer.SMS_Assigned_Sites0 
                If ($DDRFile) { Log-Append -strLogFileName $LogFileName  -strLogText ("Created DDR "+$DDRFile.FullName) }
                ELSE          { Log-Append -strLogFileName $LogFileName  -strLogText ("Failed to create DDR") }
            } 
            ELSE { Log-Append -strLogFileName $LogFileName -strLogText ("Found "+$ADComputer.Name+" in SCCM and data matches AD, nothing to do") }
        }
    }
}

    
$objDDRsToMove = @()
ForEach ( $SCCMSite in $objSMSSites ) {
    $objDDRsToMove = Get-ChildItem -Path ($DDRTempFolder+$SCCMSite.SiteCode+"*.ddr")   
    If ($objDDRsToMove.Count -gt 0) {
        ForEach ($DDRFile in $objDDRsToMove ) {
            If ( $DDRFile.Name.ToString().Substring(0,3) -eq $SCCMSite.SiteCode ) {
                $Result = Move-Item $DDRFile.FullName  ("\\"+$SiteServer+$DDRTargetFolder)  -Force 
                Log-Append -strLogFileName $LogFileName -strLogText ("Moving DDR to CAS server \\"+$SiteServer+$DDRTargetFolder+$DDRFile.Name)
            }
        }
    }
}
    
$SMSDisc = $Null
Log-Append -strLogFileName $LogFileName -strLogText ("Script finished ")