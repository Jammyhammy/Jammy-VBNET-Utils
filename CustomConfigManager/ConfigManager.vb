Imports System
Imports System.Xml.Linq
Imports System.Reflection
Imports System.Configuration


Namespace JammyUtils.XML

    ''' <summary>
    ''' Provides a means to apply settings during runtime from a settings XML.
    ''' Generate an XML first, modify values as needed, and then call Run() to apply the changes. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ConfigManager

        Private Shared ReadOnly Property GetXMLConfigPath() As String
            Get
                'Exception handling here.
                Return "C:\CBM\" & Assembly.GetExecutingAssembly().GetName().Name & ".Settings.xml"
            End Get
        End Property

        ''' <summary>
        ''' Generates a template XML file containing user/app settings and connection strings. 
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CreateConfigXML()

            Dim appConfigPath = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile
            Dim appSettings = XDocument.Load(appConfigPath)

            Dim configXML = (From key In appSettings.<configuration>
                             Select New XElement("configuration", key...<configSections>,
                                                 key...<connectionStrings>,
                                                 key...<applicationSettings>,
                                                 key...<userSettings>,
                                                 key...<appSettings>)) _
                            .Single()

            configXML.Save(GetXMLConfigPath())


        End Sub

        ''' <summary>
        ''' Overloaded. @path: specified .config file to gen from. Generates a config XML template for the specified .config file.
        ''' I.E: Solution Projects MainMenu, Core, Admin, where Core.dll.config contains the settings needing to be changed.
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CreateConfigXML(path As String)

            If Not IO.File.Exists(path) Then
                Throw New IO.FileNotFoundException("Could not locate the configuration file specified as PATH: " & path)
            End If

            If Not IO.Directory.Exists("C:\CBM") Then
                IO.Directory.CreateDirectory("C:\CBM")
            End If

            Dim savePath = "C:\CBM\" & Assembly.GetExecutingAssembly().GetName().Name & ".Settings.xml"

            Dim appSettings = XDocument.Load(path)

            Dim configXML = (From key In appSettings.<configuration>
                             Select New XElement("configuration", key...<configSections>,
                                                 key...<connectionStrings>,
                                                 key...<applicationSettings>,
                                                 key...<userSettings>,
                                                 key...<appSettings>)) _
                            .Single()

            configXML.Save(savePath)


        End Sub

        ''' <summary>
        ''' Reads in the XML file created from Generate() and applies the changes to the settings file. 
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub ApplyConfigXML()

            Dim loadPath = "C:\CBM\" & Assembly.GetExecutingAssembly().GetName().Name & ".Settings.xml"
            Dim loadConfig = XDocument.Load(loadPath).<configuration>

            Dim appConfigPath = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile
            Dim appConfig = XDocument.Load(appConfigPath).<configuration>

            Dim newConfig = From xml In appConfig.Elements.Union(loadConfig.Elements)
                            Group xml By xml.Name Into Group
                            Select Group.Last()

            Dim configfile = New XElement("configuration", newConfig)

            configfile.Save(appConfigPath)

        End Sub

    End Class

End Namespace