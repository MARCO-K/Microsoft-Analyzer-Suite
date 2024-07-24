#
# Module manifest for module 'Microsoft-Analyzer-Suite'
#
# Generated on: 07/24/2024
#
@{
	# Script module or binary module file associated with this manifest
	RootModule           = 'Microsoft-Analyzer-Suite.psm1'
	
	# Version number of this module.
	ModuleVersion        = '0.0.1'
	
	# ID used to uniquely identify this module
	GUID                 = '52c64b64-28b2-4277-8090-197438090673'
	
	# Author of this module
	Author               = 'Martin Willing'
	
	# Company or vendor of this module
	CompanyName          = ''
	
	# Copyright statement for this module
	Copyright            = '(c) 2024 Martin Willing at Lethal-Forensics (https://lethal-forensics.com/)'
	
	# Description of the functionality provided by this module
	Description          = 'A collection of PowerShell scripts for analyzing data from Microsoft 365 and Microsoft Entra ID.'
	
	# Supported PSEditions
	CompatiblePSEditions = 'Core', 'Desktop'
	
	# Minimum version of the PowerShell engine required by this module
	PowerShellVersion    = '5.1'

	# Name of the Windows PowerShell host required by this module
	# PowerShellHostName = ''

	# Minimum version of the Windows PowerShell host required by this module
	# PowerShellHostVersion = ''

	# Minimum version of Microsoft .NET Framework required by this module
	# DotNetFrameworkVersion = ''

	# Minimum version of the common language runtime (CLR) required by this module
	# CLRVersion = ''

	# Processor architecture (None, X86, Amd64) required by this module
	# ProcessorArchitecture = ''

	# Modules that must be imported into the global environment prior to importing
	# this module
	RequiredModules      = @(
		@{ ModuleName = 'PSFramework'; GUID = '8028b914-132b-431f-baa9-94a6952f21ff'; ModuleVersion = '1.10.0'; }
		@{ ModuleName = 'ImportExcel'; GUID = '60dd4136-feff-401a-ba27-a84458c57ede'; ModuleVersion = '7.8.0'; }
	)
	
	# Assemblies that must be loaded prior to importing this module
	# RequiredAssemblies = @()

	# Script files (.ps1) that are run in the caller's environment prior to importing this module.
	# ScriptsToProcess = @()

	# Type files (.ps1xml) to be loaded when importing this module
	# TypesToProcess = @()

	# Format files (.ps1xml) to be loaded when importing this module
	# FormatsToProcess = @()

	# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
	# NestedModules = @()
	
	# Functions to export from this module
	FunctionsToExport    = ''
	
	# Cmdlets to export from this module
	CmdletsToExport      = ''
	
	# Variables to export from this module
	# VariablesToExport = ''
	
	# Aliases to export from this module
	AliasesToExport      = ''
	
	# List of all modules packaged with this module
	# ModuleList = @()
	
	# List of all files packaged with this module
	# FileList = @()
	
	# Private data to pass to the module specified in ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
	PrivateData          = @{
		
		#Support for PowerShellGet galleries.
		PSData = @{
			
			# Tags applied to this module. These help with module discovery in online galleries.
			Tags         = @('Microsoft', 'M365', 'Cloud', 'Test', 'Entra', 'Azure', 'EntraID', 'incident-response', 'Microsoft-Graph')
			
			# A URL to the license for this module.
			LicenseUri   = 'https://github.com/evild3ad/Microsoft-Analyzer-Suite/blob/main/LICENSE'
			
			# A URL to the main website for this project.
			ProjectUri   = 'https://lethal-forensics.com/'
			
			# A URL to an icon representing this module.
			# IconUri = ''
			
			# ReleaseNotes of this module
			ReleaseNotes = 'https://github.com/evild3ad/Microsoft-Analyzer-Suite/releases/'
			
		} # End of PSData hashtable
		
	} # End of PrivateData hashtable

	# HelpInfo URI of this module
	HelpInfoURI          = 'https://github.com/evild3ad/Microsoft-Analyzer-Suite/wiki'
}