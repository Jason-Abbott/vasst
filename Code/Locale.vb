Imports AMP.LocaleHelper
Imports System.Threading
Imports System.Resources
Imports System.Globalization
Imports System.Reflection.Assembly

Public Class Locale
    Inherits System.Resources.ResourceManager

    Private _culture As CultureInfo

    '---COMMENT---------------------------------------------------------------
    '	initialize base object
    '
    '	Date:		Name:	Description:
    '	9/20/04     JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub New()
        MyBase.New("AMP.language", GetExecutingAssembly)
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Shadowed with default property to simplify calls, e.g. say("name")
    '
    '	Date:		Name:	Description:
    '	9/20/04     JEA		Creation
    '-------------------------------------------------------------------------
    Default Public Shadows ReadOnly Property GetString(ByVal index As String) As String
        Get
            Return MyBase.GetString(index)
        End Get
    End Property

    '---COMMENT---------------------------------------------------------------
    '	Build and return a drop down list with all cultures
    '
    '	Date:		Name:	Description:
    '	9/20/04     JEA		Creation
    '-------------------------------------------------------------------------
    Public Function CreateCultureDropDown() As System.Web.UI.WebControls.DropDownList
        Dim dropDownList As System.Web.UI.WebControls.DropDownList

        For Each _culture In CultureInfo.GetCultures(CultureTypes.SpecificCultures)
            Dim item As New System.Web.UI.WebControls.ListItem
            item.Value = _culture.Name
            item.Text = _culture.DisplayName
            dropDownList.Items.Add(item)
        Next
        Return dropDownList
    End Function

    '---COMMENT---------------------------------------------------------------
    '	set the thread culture to given culture
    '   Culture is used for date/time localization
    '   UICulture is used to load localized resources
    '
    '	Date:		Name:	Description:
    '	9/20/04     JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub SetCulture(ByVal culture As CultureInfo)
        With Thread.CurrentThread
            .CurrentCulture = culture
            .CurrentUICulture = culture
        End With
    End Sub

    ' overload to get culture from string of culture, such as from drop-down
    Public Sub SetCulture(ByVal culture As String)
        SetCulture(New CultureInfo(culture))
    End Sub

End Class


Public Class LocaleHelper
    '---COMMENT---------------------------------------------------------------
    '   Return namespace of language resource file; assumes file in project root
    '   named language.resx and that the assembly file name matches the namespace
    '	This method must be in a separate class since it's called from the constructor
    '   of the main class above; you can't call a method in a class still under construction
    '
    '	Date:		Name:	Description:
    '	9/20/04     JEA		Creation
    '-------------------------------------------------------------------------
    Shared Function LanguageNameSpace(ByVal assemblyName As String) As String
        Dim attributes As String() = assemblyName.Split(","c)
        Return attributes(0).TrimEnd + ".language"
    End Function
End Class
