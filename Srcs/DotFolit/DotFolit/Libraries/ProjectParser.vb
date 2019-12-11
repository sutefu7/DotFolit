Imports System.Data
Imports System.IO


Public Class ProjectParser

    Public Function GetAssemblyFile(projectFile As String) As String

        Dim prjDir = Path.GetDirectoryName(projectFile)
        Dim element = XElement.Load(projectFile)
        Dim ns As XNamespace = element.Attribute("xmlns").Value
        Dim assemblyName = element.Descendants(ns + "AssemblyName").FirstOrDefault().Value
        Dim outputType = element.Descendants(ns + "OutputType").FirstOrDefault().Value
        Dim extension = ".exe"

        If outputType = "Library" Then
            extension = ".dll"
        End If

        ' bin/Debug, bin/Release, bin/x86/Debug, ... 等の組み合わせがある
        Dim conditionElements = From x In element.Descendants(ns + "PropertyGroup")
                                Where x.HasAttributes AndAlso x.Attributes("Condition").Any()
                                Where x.Attribute("Condition").Value.Contains("$(Configuration)|$(Platform)")

        Dim outputPaths = New List(Of String)
        For Each conditionElement In conditionElements

            Dim outputPath = conditionElement.Descendants(ns + "OutputPath").FirstOrDefault().Value
            outputPaths.Add(outputPath)

        Next

        ' ビルド構成とプラットフォームの組み合わせのうち、以下の優先度で、ビルド済みアセンブリファイルのパスを返却する
        Dim checkPaths = New List(Of String) From {"bin\Debug", "bin\x64\Debug", "bin\x86\Debug", "bin\Release", "bin\x64\Release", "bin\x86\Release"}
        For Each checkPath In checkPaths

            If outputPaths.Any(Function(x) x.Contains(checkPath)) Then

                Dim outputPath = outputPaths.FirstOrDefault(Function(x) x.Contains(checkPath))
                outputPaths.Remove(outputPath)

                Dim assemblyFile = Path.Combine(prjDir, outputPath, $"{assemblyName}{extension}")
                If File.Exists(assemblyFile) Then
                    Return assemblyFile
                End If

            End If

        Next

        ' その他あれば、残りは順不同
        For Each outputPath In outputPaths

            Dim assemblyFile = Path.Combine(prjDir, outputPath, $"{assemblyName}{extension}")
            If File.Exists(assemblyFile) Then
                Return assemblyFile
            End If

        Next

        Return String.Empty

    End Function

    Public Function GetDocumentFile(projectFile As String) As String

        Dim prjDir = Path.GetDirectoryName(projectFile)
        Dim element = XElement.Load(projectFile)
        Dim ns As XNamespace = element.Attribute("xmlns").Value

        ' bin/Debug, bin/Release, bin/x86/Debug, ... 等の組み合わせがある
        Dim conditionElements = From x In element.Descendants(ns + "PropertyGroup")
                                Where x.HasAttributes AndAlso x.Attributes("Condition").Any()
                                Where x.Attribute("Condition").Value.Contains("$(Configuration)|$(Platform)")

        Dim items = New List(Of Tuple(Of String, String))
        For Each conditionElement In conditionElements

            Dim outputPath = conditionElement.Descendants(ns + "OutputPath").FirstOrDefault().Value
            Dim documentName = conditionElement.Descendants(ns + "DocumentationFile").FirstOrDefault().Value
            items.Add(Tuple.Create(outputPath, documentName))

        Next

        ' ビルド構成とプラットフォームの組み合わせのうち、以下の優先度で、ビルド済みアセンブリファイルのパスを返却する
        Dim checkPaths = New List(Of String) From {"bin\Debug", "bin\x64\Debug", "bin\x86\Debug", "bin\Release", "bin\x64\Release", "bin\x86\Release"}
        For Each checkPath In checkPaths

            If items.Any(Function(x) x.Item1.Contains(checkPath)) Then

                Dim item = items.FirstOrDefault(Function(x) x.Item1.Contains(checkPath))
                items.Remove(item)

                Dim documentFile = Path.Combine(prjDir, item.Item1, item.Item2)
                If File.Exists(documentFile) Then
                    Return documentFile
                End If

            End If

        Next

        ' その他あれば、残りは順不同
        For Each item In items

            Dim documentFile = Path.Combine(prjDir, item.Item1, item.Item2)
            If File.Exists(documentFile) Then
                Return documentFile
            End If

        Next

        Return String.Empty

    End Function

    Public Iterator Function GetImportNamespaceNames(projectFile As String) As IEnumerable(Of String)

        Dim prjDir = Path.GetDirectoryName(projectFile)
        Dim element = XElement.Load(projectFile)
        Dim ns As XNamespace = element.Attribute("xmlns").Value
        Dim importElements = From x In element.Descendants(ns + "Import")
                             Where x.HasAttributes AndAlso x.Attributes("Include").Any()

        For Each importElement In importElements

            Dim importName = importElement.Attribute("Include").Value
            Yield importName

        Next

    End Function

    Public Iterator Function GetReferenceAssemblyNames(projectFile As String) As IEnumerable(Of String)

        Dim prjDir = Path.GetDirectoryName(projectFile)
        Dim element = XElement.Load(projectFile)
        Dim ns As XNamespace = element.Attribute("xmlns").Value

        ' テストプロジェクトの場合？、参照 dll の読み込み有無を、条件分岐している場合がある
        ' この場合、記載の評価通りに評価して、参照するかどうか判定してから返却する
        Dim isTestProject = element.Descendants(ns + "TestProjectType").Any()
        Dim isCodedUITest = False
        Dim vsVersion = String.Empty
        Dim targetFwVersion = String.Empty

        If isTestProject Then

            Dim tmpElement = element.Descendants(ns + "IsCodedUITest").FirstOrDefault()
            isCodedUITest = CBool(tmpElement.Value)

            tmpElement = element.Descendants(ns + "VisualStudioVersion").FirstOrDefault()
            vsVersion = tmpElement.Value

            tmpElement = element.Descendants(ns + "TargetFrameworkVersion").FirstOrDefault()
            targetFwVersion = tmpElement.Value

        End If

        ' Reference タグを取得
        Dim referenceElements = From x In element.Descendants(ns + "Reference")
                                Where x.HasAttributes AndAlso x.Attributes("Include").Any()

        For Each referenceElement In referenceElements

            ' 条件がある場合、評価して返却するかどうかを、事前に判定する
            If referenceElement.Parent.Parent.Name = "When" OrElse referenceElement.Parent.Parent.Name = "Otherwise" Then

                Dim chooseElement = referenceElement.Parent.Parent.Parent
                Dim whenElement = chooseElement.Descendants(ns + "When").FirstOrDefault()
                Dim condition = whenElement.Attribute("Condition").Value

                ' ('$(VisualStudioVersion)' == '10.0' or '$(VisualStudioVersion)' == '') and '$(TargetFrameworkVersion)' == 'v3.5'
                ' ('10.0' == '10.0' or '10.0' == '') and 'v4.6' == 'v3.5'
                ' ('10.0' = '10.0' or '10.0' = '') and 'v4.6' = 'v3.5'

                condition = condition.Replace("$(VisualStudioVersion)", vsVersion)
                condition = condition.Replace("$(TargetFrameworkVersion)", targetFwVersion)
                condition = condition.Replace("$(IsCodedUITest)", isCodedUITest.ToString())
                condition = condition.Replace("(", " ( ")
                condition = condition.Replace(")", " ) ")
                condition = condition.Replace(" == ", " = ")
                condition = condition.Replace(" != ", " <> ")
                condition = condition.Replace(ControlChars.Quote, String.Empty)

                ' 文字列を読み取って、評価できる機能を利用する（自前解析は断念）
                Dim dt = New DataTable
                Dim result = dt.Compute(condition, String.Empty)
                Dim isTrueOfWhenCondition = CBool(result)

                If referenceElement.Parent.Parent.Name = "When" Then

                    ' When 要素タグの下ならば、True だった = この参照dllを読み込む。なので、False の場合、はじく
                    If Not isTrueOfWhenCondition Then
                        Continue For
                    End If

                Else

                    ' Otherwise 要素タグの下ならば、When の評価結果が False だった = こちら（Otherwise）の参照dllを読み込む。なので、True の場合、はじく
                    If isTrueOfWhenCondition Then
                        Continue For
                    End If

                End If

            End If

            Dim referenceName = referenceElement.Attribute("Include").Value
            If referenceName.Contains(", Version=") Then
                referenceName = referenceName.Substring(0, referenceName.IndexOf(", Version="))
            End If

            Yield referenceName

        Next

        ' ProjectReference タグを取得
        referenceElements = From x In element.Descendants(ns + "ProjectReference")
                            Where x.HasAttributes AndAlso x.Attributes("Include").Any()

        For Each referenceElement In referenceElements

            Dim projectName = referenceElement.Descendants(ns + "Name").FirstOrDefault()
            Yield projectName.Value

        Next

    End Function

    Public Iterator Function GetSourceFiles(projectFile As String) As IEnumerable(Of String)

        Dim prjDir = Path.GetDirectoryName(projectFile)
        Dim element = XElement.Load(projectFile)
        Dim ns As XNamespace = element.Attribute("xmlns").Value
        Dim compileElements = From x In element.Descendants(ns + "Compile")
                              Where x.HasAttributes AndAlso x.Attributes("Include").Any()

        For Each compileElement In compileElements

            Dim sourceName = compileElement.Attribute("Include").Value
            If sourceName.Contains("My Project") Then
                Continue For
            End If

            Dim sourceFile = Path.Combine(prjDir, sourceName)
            Yield sourceFile

        Next

    End Function

    Public Iterator Function GetSourceFilesWithDependentUpon(projectFile As String) As IEnumerable(Of Tuple(Of String, String))

        Dim prjDir = Path.GetDirectoryName(projectFile)
        Dim element = XElement.Load(projectFile)
        Dim ns As XNamespace = element.Attribute("xmlns").Value
        Dim compileElements = From x In element.Descendants(ns + "Compile")
                              Where x.HasAttributes AndAlso x.Attributes("Include").Any()

        For Each compileElement In compileElements

            Dim sourceName = compileElement.Attribute("Include").Value
            If sourceName.Contains("My Project") Then
                Continue For
            End If

            Dim dependentUpon = String.Empty
            If compileElement.Descendants(ns + "DependentUpon").Any() Then

                ' 対応するソースがフォルダに入っていた場合でも、依存ソースの記載がファイル名だけになっているため、
                ' 同じサブフォルダになるように追加する
                Dim subDir = String.Empty
                If sourceName.Contains("\") Then
                    subDir = sourceName.Substring(0, sourceName.LastIndexOf("\") + 1)
                End If

                dependentUpon = compileElement.Descendants(ns + "DependentUpon").FirstOrDefault().Value
                dependentUpon = Path.Combine(prjDir, $"{subDir}{dependentUpon}")

            End If

            Dim sourceFile = Path.Combine(prjDir, sourceName)
            Dim dependentUponFile = dependentUpon

            Yield Tuple.Create(sourceFile, dependentUponFile)

        Next

    End Function

    Public Function GetRootNamespace(projectFile As String) As String

        Dim prjDir = Path.GetDirectoryName(projectFile)
        Dim element = XElement.Load(projectFile)
        Dim ns As XNamespace = element.Attribute("xmlns").Value
        Dim rootNamespace = element.Descendants(ns + "RootNamespace").FirstOrDefault().Value

        Return rootNamespace

    End Function

End Class
