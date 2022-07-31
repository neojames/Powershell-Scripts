# ADGetUsers
# James Bolton 2022
# james@neojames.me


<#
.Synopsis
	Output group members.
.DESCRIPTION
	Simply output members of a group in a nice(ish) HTML file.
.EXAMPLE
	TODO: Fill out later.
.PARAMETER Path
	Path to output folder. Defaults to current working directory.
.PARAMETER Overwrite
	If Output exists should I overwrite?
.PARAMETER Online
	Will we need to query Distribution Lists?
.PARAMETER Credential
	Office 365 username
.LINK
	TODO: Add github
#>

 param (
	[Parameter()]
	[string]$Path = $($(Get-Location | Get-Member -Force -Static -Type Method op_*) + "output.html"),
	[Parameter()]
	[switch]$Overwrite,
	[Parameter()]
	[switch]$Online,
	[Parameter()]
	[string]$Credential
 )

 if ($Online -eq $true) {
	try{
		Import-Module ExchangeOnlineManagement
	}
	catch{
		Write-Host "Please install ExchangeOnlineManagement with Install-Module -Name ExchangeOnlineManagement as an Administrator."
		Exit 1
	}
	Connect-ExchangeOnline -UserPrincipalName $Credential -ShowBanner:$false
}

# Check if Output Exists

if (Test-Path -Path $path) {
	if ($overwrite -eq $false) {
		Write-Host "Error: That file already exists, if you would like to overwrite please run with the -Overwrite switch"
		exit 1
	}
}

# Create HTML Skellington
$head=@"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>AD Groups</title>
	<script type="text/javascript">
	window.onload = function () {
    var toc = "";
    var level = 0;

    document.getElementById("contents").innerHTML =
        document.getElementById("contents").innerHTML.replace(
            /<h([\d])>([^<]+)<\/h([\d])>/gi,
            function (str, openLevel, titleText, closeLevel) {
                if (openLevel != closeLevel) {
                    return str;
                }

                if (openLevel > level) {
                    toc += (new Array(openLevel - level + 1)).join("<ul>");
                } else if (openLevel < level) {
                    toc += (new Array(level - openLevel + 1)).join("</ul>");
                }

                level = parseInt(openLevel);

                var anchor = titleText.replace(/ /g, "_");
                toc += "<li><a href=\"#" + anchor + "\">" + titleText
                    + "</a></li>";

                return "<h" + openLevel + "><a name=\"" + anchor + "\">"
                    + titleText + "</a></h" + closeLevel + ">";
            }
        );

    if (level) {
        toc += (new Array(level + 1)).join("</ul>");
    }

    document.getElementById("toc").innerHTML += toc;
};
	</script>
</head>
<body>
	<h1>Active Directory Group Membership</h1>
	<div id="toc">
		<h2>Table of Contents</h2>
	</div>
	<div id="contents">
"@
$head | Out-File -FilePath $path

# Create a list of every AD Group
$groups = Get-ADGroup -Filter * | select name, GroupCategory

# Generate table for each group
foreach ($i in $groups) {
	$name = $($i | select -expandproperty name)
	$type = $($i | select -expandproperty GroupCategory)
	$output ="<h2>" + $name + "</h2>"
	$output | Out-File -FilePath $path -Append

	$output ="<p>Group Type:" + $type + "<p>"
	$output | Out-File -FilePath $path -Append

	if (($type -eq "Distribution") -and ($Online -eq $true)) {
		$table = Get-DistributionGroupMember -ResultSize Unlimited -Identity $name | select name | ConvertTo-Html
	}
	else {
		$table = Get-ADGroupMember -Identity $name | select name | ConvertTo-Html
	}

	$table = $table.replace('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> <html xmlns="http://www.w3.org/1999/xhtml"> <head> <title>HTML TABLE</title> </head><body> ',"")
	$table = $table.replace('</body></html>',"")
	$table = $table.replace('<th>*</th>', "")
	$table | Out-File -FilePath $path -Append
}

# Do the same for Office 365 Distribution lists if online
if ($Online -eq $true) {
	$onlineGroups = Get-DistributionGroup -Filter * | select DisplayName
	foreach ($i in $onlineGroups) {
		$name = $($i | select -expandproperty DisplayName)
		$type = "Office 365 Distribution Group"
		$output ="<h2>" + $name + "</h2>"
		$output | Out-File -FilePath $path -Append

		$output ="<p>Group Type:" + $type + "<p>"
		$output | Out-File -FilePath $path -Append

		$table = Get-DistributionGroupMember -ResultSize Unlimited -Identity $name | select name | ConvertTo-Html
		$table = $table.replace('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> <html xmlns="http://www.w3.org/1999/xhtml"> <head> <title>HTML TABLE</title> </head><body> ',"")
		$table = $table.replace('</body></html>',"")
		$table = $table.replace('<th>*</th>', "")
		$table | Out-File -FilePath $path -Append 
	}
	Disconnect-ExchangeOnline -Confirm:$false
}

# Close out the file
$foot=@"
	</div>
	</body>
</html
"@
$foot | Out-File -FilePath $path -Append