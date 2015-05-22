Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' Les informations générales relatives à un assembly dépendent de 
' l'ensemble d'attributs suivant. Changez les valeurs de ces attributs pour modifier les informations
' associées à un assembly.

' Passez en revue les valeurs des attributs de l'assembly

<Assembly: AssemblyTitle("TestBouton")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("TestBouton")> 
<Assembly: AssemblyCopyright("Copyright ©  2015")> 
<Assembly: AssemblyTrademark("")> 

' L'affectation de la valeur false à ComVisible rend les types invisibles dans cet assembly 
' aux composants COM.  Si vous devez accéder à un type dans cet assembly à partir de 
' COM, affectez la valeur true à l'attribut ComVisible sur ce type.
<Assembly: ComVisible(False)>

'Le GUID suivant est pour l'ID de la typelib si ce projet est exposé à COM
<Assembly: Guid("ab5e4cd6-745e-4e6d-9452-c8672ec67397")> 

' Les informations de version pour un assembly se composent des quatre valeurs suivantes :
'
'      Version principale
'      Version secondaire 
'      Numéro de build
'      Révision
'
' Vous pouvez spécifier toutes les valeurs ou indiquer les numéros de build et de révision par défaut 
' en utilisant '*', comme indiqué ci-dessous :
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
