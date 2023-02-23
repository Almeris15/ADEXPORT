#NAME : ExportAD.ps1
#Author : Almeris15
#Description : Script pour exporter la liste des utilisateurs, les groupes et les membres des groupes au format CSV pour une liste d'OU fourni en paramètre sur un AD
#  
# Changelog:
# 1.0.0 - Initial release

# Import the Active Directory module
Import-Module ActiveDirectory

# ----- VARIABLES -----
# Define Ad Name
$Adname = $args[0]
$AD_OU = $Adname.Split('.')

# Define the list of OUs to search
$ous = ""
foreach ($i in $AD_OU) {
	$ous += "DC=" + $i + ","
}
$ous = $ous.TrimEnd(',')

# Define Date
$date = Get-Date -Format "yyyyMMdd"

# Define CSV exit
# The folders must be created on  the server where the script will be executed
$Path = "C:\SAFEO\FILES\"

# Name CSV
$Name_CSV_Groups = $Path + "ADgroups_" + $ADname + "_" + $date + ".csv"
$Name_CSV_User = $Path + "ADmembers_" + $ADname + "_" + $date + ".csv"
$Name_CSV_GroupMembership = $Path + "ADmembership_" + $ADname + "_" + $date + ".csv"
# ----- FIN VARIABLES ------

# Initialize arrays to store the group and member information
$groups = @()
$members = @()
$groupMembership = @()

# Get all of the groups in the specified OUs
$groups += Get-ADGroup -Filter * -SearchBase $ous


# Loop through each group
foreach ($group in $groups) {
	# Get the members of the group
	$groupMembers = Get-ADGroupMember $group

	# Add the group to the list of groups
	$groupsInfo = New-Object -TypeName PSObject -Property @{
	        GroupName = $group.Name
	}
	$groups += $groupsInfo

	# Loop through each member of the group
	foreach ($member in $groupMembers) {
		# Add the member to the list of members
		$membersInfo = New-Object -TypeName PSObject -Property @{
			MemberSamAccountName = $member.SamAccountName
		}
		$members += $membersInfo
		
		# Add the group membership to the list of group memberships
		$groupMembershipInfo = New-Object -TypeName PSObject -Property @{
			GroupName = $group.Name
			MemberSamAccountName = $member.SamAccountName
		}
		$groupMembership += $groupMembershipInfo
	}
}

# Remove duplicates from the groups array
$uniqueGroups = $groups | Select-Object -Property GroupName -Unique

# Remove duplicates from the members array
$uniqueMembers = $members | Select-Object -Property MemberSamAccountName -Unique

# Export the group information to a CSV file
$uniqueGroups | Export-Csv -Path $Name_CSV_Groups -NoTypeInformation

# Export the member information to a CSV file
$uniqueMembers | Export-Csv -Path $Name_CSV_User -NoTypeInformation

# Export the group membership information to a CSV file
$groupMembership | Export-Csv -Path $Name_CSV_GroupMembership -NoTypeInformation

# Display a message indicating that the export is complete
Write-Host "Export complete!"
