Imports DevExpress.Docs.Presentation
Imports DevExpress.Drawing


Namespace ExtractPresentationContent

	Public Class Program
		Public Shared Sub Main(ByVal underscore() As String)
			Dim presentation As New Presentation(File.ReadAllBytes("data\Sample.pptx"))

			Console.WriteLine("Choose what to extract:")
			Console.WriteLine("1. Slide text")
			Console.WriteLine("2. Text from all slides")
			Console.WriteLine("3. Text from a specific paragraph")
			Console.WriteLine("4. Slide notes")
			Console.WriteLine("5. Slide note body text")
			Console.WriteLine("6. Notes from all slides")
			Console.WriteLine("7. Slide pictures")
			Console.WriteLine("8. Pictures from all slides")

			Console.Write("Enter option number: ")
			Dim input As String = Console.ReadLine()

			Select Case input
				Case "1"
					SlideText(presentation, 1)
				Case "2"
					SlidesText(presentation)
				Case "3"
					ParagraphText(presentation)
				Case "4"
					SlideNoteText(presentation,0)
				Case "5"
					SlideNoteBodyText(presentation, 0)
				Case "6"
					SlidesNoteText(presentation)
				Case "7"
					SaveSlidePictures(presentation, 1)
				Case "8"
					SaveSlidesPictures(presentation)
				Case Else
					Console.WriteLine("Invalid option.")
			End Select
		End Sub

		Public Shared Function SlideText(ByVal presentation As Presentation, ByVal slideNumber As Integer) As String
			' Extract text from the second slide 
			Dim slidesText As String = ""

'			#Region "Sort shapes"
			Dim sortedShapes = presentation.Slides(slideNumber).Shapes.Where(Function(shape) TypeOf shape Is Shape AndAlso CType(shape, Shape).TextArea IsNot Nothing).OrderBy(Function(shape) shape.Y).ThenBy(Function(shape) shape.X)
'			#End Region

			For Each shape In sortedShapes
				Dim tempVar As Boolean = TypeOf shape Is Shape
				Dim textShape As Shape = If(tempVar, CType(shape, Shape), Nothing)
				If tempVar Then
					Dim shapeText As String = textShape.TextArea.Text

'					#Region "Filter shapes"
					If textShape.PlaceholderSettings?.Type = PlaceholderType.SlideNumber OrElse String.IsNullOrWhiteSpace(shapeText) Then
						Continue For
					End If
'					#End Region

					slidesText &= shapeText & vbCrLf
				End If
			Next shape
			Return slidesText
		End Function

		Public Shared Function SlidesText(ByVal presentation As Presentation) As String
			' Extract text from all slides
			Dim slideNumber As Integer = 0
			Dim slidesText_Conflict As String = ""
			For Each slide In presentation.Slides
'				#Region "Sort shapes"
				Dim sortedShapes = slide.Shapes.Where(Function(shape) TypeOf shape Is Shape AndAlso CType(shape, Shape).TextArea IsNot Nothing).OrderBy(Function(shape) shape.Y).ThenBy(Function(shape) shape.X)
'				#End Region

				For Each shape In sortedShapes
					Dim tempVar As Boolean = TypeOf shape Is Shape
					Dim textShape As Shape = If(tempVar, CType(shape, Shape), Nothing)
					If tempVar Then
						Dim shapeText As String = textShape.TextArea.Text

'						#Region "Filter shapes"
						If textShape.PlaceholderSettings?.Type = PlaceholderType.SlideNumber OrElse String.IsNullOrWhiteSpace(shapeText) Then
							Continue For
						End If
'						#End Region

						slidesText_Conflict &= shapeText & vbCrLf
					End If
				Next shape
				slideNumber += 1
			Next slide
			Return slidesText_Conflict
		End Function

		Public Shared Function ParagraphText(ByVal presentation As Presentation) As String
			'Extract text from the specific shape's paragraph
			Dim paraText As String = ""
			Dim shape As Shape = presentation.Slides(0).Shapes.Find(Of Shape)(Function(s) s.Name = "TextBox 3")
			paraText = shape.TextArea.Paragraphs(1).Text
			Return paraText
		End Function

		Public Shared Function SlideNoteText(ByVal presentation As Presentation, ByVal slideNumber As Integer) As String
			' Extract note text from the first slide 
			Dim slideNoteText_Conflict As String = ""
			For Each noteShape In presentation.Slides(slideNumber).Notes.Shapes
				Dim tempVar As Boolean = TypeOf noteShape Is Shape
				Dim textNoteShape As Shape = If(tempVar, CType(noteShape, Shape), Nothing)
				If tempVar Then
					Dim noteShapeText As String = textNoteShape.TextArea.Text
'					#Region "Filter shapes"
					If textNoteShape.PlaceholderSettings?.Type = PlaceholderType.SlideNumber OrElse String.IsNullOrWhiteSpace(noteShapeText) Then
						Continue For
					End If
'					#End Region
					slideNoteText_Conflict &= noteShapeText & vbCrLf
				End If
			Next noteShape
			Return slideNoteText_Conflict
		End Function
		Public Shared Function SlideNoteBodyText(ByVal presentation As Presentation, ByVal slideNumber As Integer) As String
			Dim notesText As String = ""
			For Each noteShape In presentation.Slides(slideNumber).Notes.Shapes
				Dim tempVar As Boolean = TypeOf noteShape Is Shape
				Dim textNoteShape As Shape = If(tempVar, CType(noteShape, Shape), Nothing)
				If tempVar AndAlso textNoteShape.PlaceholderSettings.Type = PlaceholderType.Body Then
					Dim noteShapeText As String = textNoteShape.TextArea.Text
'					#Region "Filter shapes"
					If textNoteShape.PlaceholderSettings?.Type = PlaceholderType.SlideNumber OrElse String.IsNullOrWhiteSpace(noteShapeText) Then
						Continue For
					End If
					notesText &= noteShapeText & vbCrLf
'					#End Region
				End If
			Next noteShape
			Return notesText
		End Function
		Public Shared Function SlidesNoteText(ByVal presentation As Presentation) As String
			Dim notesText As String = ""
			For Each slide In presentation.Slides
				If slide.Notes.Shapes.Any() Then
'					#Region "Sort shapes"
					Dim sortedNoteShapes = slide.Notes.Shapes.Where(Function(shape) TypeOf shape Is Shape AndAlso CType(shape, Shape).TextArea IsNot Nothing).OrderBy(Function(shape) shape.Y).ThenBy(Function(shape) shape.X)
'					#End Region
					For Each noteShape In sortedNoteShapes
						Dim tempVar As Boolean = TypeOf noteShape Is Shape
						Dim textNoteShape As Shape = If(tempVar, CType(noteShape, Shape), Nothing)
						If tempVar Then
							Dim noteShapeText As String = textNoteShape.TextArea.Text
'							#Region "Filter shapes"
							If textNoteShape.PlaceholderSettings?.Type = PlaceholderType.SlideNumber OrElse String.IsNullOrWhiteSpace(noteShapeText) Then
								Continue For
							End If
							notesText &= noteShapeText & vbCrLf
'							#End Region
						End If
					Next noteShape
				End If
			Next slide
			Return notesText
		End Function
		Public Shared Sub SaveSlidePictures(ByVal presentation As Presentation, ByVal slideNumber As Integer)
			' Extract pictures from the second slide 
'			#Region "Sort shapes"
			Dim sortedShapes = presentation.Slides(slideNumber).Shapes.Where(Function(shape) TypeOf shape Is PictureShape).OrderBy(Function(shape) shape.Y).ThenBy(Function(shape) shape.X)
'			#End Region

			Dim imageCount As Byte = 0
			For Each pictureShape As PictureShape In sortedShapes
				pictureShape.Image.Save("Slide2_Picture" & imageCount & ".png", DXImageFormat.Png)
				imageCount += 1
			Next pictureShape

		End Sub
		Public Shared Sub SaveSlidesPictures(ByVal presentation As Presentation)
			' Extract pictures from all slides
			Dim slideNumber As Integer = 0
			For Each slide In presentation.Slides
'				#Region "Sort shapes"
				Dim sortedShapes = slide.Shapes.Where(Function(shape) TypeOf shape Is PictureShape).OrderBy(Function(shape) shape.Y).ThenBy(Function(shape) shape.X)
'				#End Region

				Dim imageCount As Byte = 0
				For Each pictureShape As PictureShape In sortedShapes
					pictureShape.Image.Save("Slide" & slideNumber & "_Picture" & imageCount & ".png", DXImageFormat.Png)
					imageCount += 1
				Next pictureShape
				slideNumber += 1
			Next slide
		End Sub
	End Class
End Namespace
