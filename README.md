<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/1027678744/25.1.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1301588)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# Presentation Document API – Extract Presentation Images, Notes, and Pictures

Code samples within this repository extract content from presentation files (including shape text, images, and slide notes).

Use the following collections to obtain slides and associated content:

* [Presentation.Slides](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Docs.Presentation.Presentation.Slides) - stores all slides within the presentation. 
* [Slide.Shapes](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Docs.Presentation.SlideBase.Shapes) - stores shapes within a slide. Use shapes to access content such as text, images, figures, etc. 

## Implementation Details

This repository uses the following techniques to correctly obtain slide content:

* The order of elements in the [Slide.Shapes](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Docs.Presentation.SlideBase.Shapes) collection does not necessarily match visual order. To process shapes top to bottom or left to right, sort them by coordinates. Examples in this repository use the [System.Linq](https://learn.microsoft.com/en-us/dotnet/api/system.linq?view=net-9.0) namespace to sort elements.
    
    **Code to review**: the **Sort shapes** region.

* You can identify and skip certain shape types. For example, the [Slide.Shapes](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Docs.Presentation.SlideBase.Shapes) collection includes a slide number placeholder shape. Use the [CommonShape.PlaceholderSettings.Type](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Docs.Presentation.PlaceholderSettings.Type) property to identify and skip this shape. 

   **Code to review**: the **Filter shapes** region.

* You can obtain only shapes that contain text. Check that the [TextArea.Text](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Docs.Presentation.TextArea.Text) property is not an empty string.

    **Code to review**: the **Filter shapes** region.

## File to Review

[Program.cs](./CS/Program.cs) ([Program.vb](./VB/Program.vb))


## Documentation

Refer to the following help topic for image/extraction results: [Extract Presentation Content](https://docs.devexpress.com/OfficeFileAPI/405430/presentation-api/extract-presentation-content).

<!-- ## More Examples -->

<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=presentation-api-extract-content&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=presentation-api-extract-content&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->


