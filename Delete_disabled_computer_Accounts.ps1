################################################################################
##			Script to Delete all disabled computer accounts					  ##
##Created by David Hahn - 9/10/2010											  ##
##Found how to use get-qadcomputer to return only disabled objects from this  ##
##blog:																		  ##
##http://dmitrysotnikov.wordpress.com/2010/04/12/get-enabled-or-disabled-computer-accounts/
################################################################################
$searchroot = 'DC=contoso,DC=com'
#Delete all disabled computer accounts
Get-QADComputer -ldapFilter ‘(userAccountControl:1.2.840.113556.1.4.803:=2)’ -SearchRoot $searchroot |
Remove-QADObject -WhatIf

