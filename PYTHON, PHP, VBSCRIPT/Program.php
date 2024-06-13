<?php
  // Create ComHelper object.
  $comHelper = new Com("GemBox.Document.ComHelper", null, CP_UTF8);
  // If using the Professional version, put your serial key below.
  $comHelper->ComSetLicense("FREE-LIMITED-KEY");
  
  // Read Word document.
  $document = $comHelper->Load(getcwd() . "\ComTemplate.docx");
  
  // Find and replace text.
  $document->Content->Replace("PLACEHOLDER1", "Sample Value 1");
  $document->Content->Replace("PLACEHOLDER2", "Sample Value 2");
  $document->Content->Replace("PLACEHOLDER3", "Sample Value 3");
  
  // Execute mail merge process.
  $source = new Com("System.Collections.Hashtable", null, CP_UTF8);
  $source->Add("Name", "John");
  $source->Add("Surname", "Doe");
  $source->Add("Age", 30);
  $document->MailMerge->Execute($source);
  
  // Modify bookmarks content.
  $document->Bookmarks->Item("Bookmark1")->GetContent(true)->LoadText("Sample Content 1.");
  $document->Bookmarks->Item("Bookmark2")->GetContent(true)->LoadText("Sample Content 2.");
  
  // Write document as PDF.
  $document->Save(getcwd() . "\ComExample.pdf");
?>