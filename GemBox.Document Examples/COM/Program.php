// Create ComHelper object and set license. 
// NOTE: If you're using a Professional version you'll need to put your serial key below.
$comHelper = new Com("GemBox.Document.ComHelper", null, CP_UTF8);
$comHelper->ComSetLicense("FREE-LIMITED-KEY");

// Load input document.
$document = $comHelper->Load(getcwd() . "\ComTemplate.docx");

// Find and replace
$document->Content->Replace("PLACEHOLDER1", "Sample Value 1");
$document->Content->Replace("PLACEHOLDER2", "Sample Value 2");
$document->Content->Replace("PLACEHOLDER3", "Sample Value 3");

// Mail merge
$source = new Com("System.Collections.Hashtable", null, CP_UTF8);
$source->Add("Name", "John");
$source->Add("Surname", "Doe");
$source->Add("Age", 30);
$document->MailMerge->Execute($source);

// Modify bookmarks
$document->Bookmarks->Item("Bookmark1")->GetContent(true)->LoadText("Sample Content 1.");
$document->Bookmarks->Item("Bookmark2")->GetContent(true)->LoadText("Sample Content 2.");

// Get output path and save document as PDF document.
$path = getcwd() . "\ComExample.pdf";

$document->Save($path);
echo("Document saved as '" . $path . "'");