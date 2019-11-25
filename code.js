// Created for Challenger Sports by Joe Hays July 2018
var ss=SpreadsheetApp.getActiveSpreadsheet();
var sheet=ss.getActiveSheet();
var ui=SpreadsheetApp.getUi();
var data=sheet.getDataRange().getValues();
var range=sheet.getDataRange();
var values=range.getValues();
var count=1;
var Term;

// This calls the onOpen function when the Add-on is first installed
function onInstall(e){
    onOpen(e);
}

// The onOpen function is executed automatically every time a Spreadsheet is loaded
function onOpen(e){
    var menu=SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp or FormApp.
    menu.addItem('Run','Input');
    menu.addToUi();
}

//Prompts user for some initial inputs
function Input(){
    var htmlOutput=HtmlService
        .createHtmlOutputFromFile('sidebar')
        .setTitle('Mail Merge')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    ui.showSidebar(htmlOutput);
}

// This creates a folder with the current time and date and appends a copy of the specified file to that folder and makes edits to the file
function CreateFolderAndDoc(Title, cloneID){
    if (cloneID !=  ''){
        //Find time and date
        var TZ=CalendarApp.getDefaultCalendar().getTimeZone();
        var currentTime=Utilities.formatDate(new Date(), TZ, 'MM/dd/yyyy-HH:mm');
    
        // Error checking for empty title variable
        var FolderTitle='';
        if(Title !=  ''){
            FolderTitle=Title+'-';
        }
      
        // Find and define the columns --- var <ANY_NAME>=findHeader('<TITLE_OF_COLUMN>');
        var ClubNameIndex=findHeader('Club Name');
        var ClubLogoIndex=findHeader('Club Logo');
        var ContactNameIndex=findHeader('Contact Name');
        var ContactNumberIndex=findHeader('Contact Number');
        var FooterNumberIndex=findHeader('Footer Number');
    
        // Create Folder using date and time found above
        DriveApp.createFolder(FolderTitle+currentTime);
        var folder=DriveApp.getFoldersByName(FolderTitle+currentTime).next().getId();
        var folderID=DriveApp.getFolderById(folder);
        folderID.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT); //Set Access to parent folder - ANYONE on challengersports.com can EDIT
        
        // Gets Values of cells, creates document using the Club Name and appends values to variables
        for (var i=1; i<values.length; i++){
            // Get Values of cells to fill variables --- var <ANY_NAME>=sheet.getRange(i+1,<COLUMN_DEFINITION_VAR>).getValue();
            var ClubNameValue=sheet.getRange(i+1,ClubNameIndex).getValue();
            var ClubLogoValue=sheet.getRange(i+1,ClubLogoIndex).getValue();
            var driveImage=DriveApp.getFileById(ClubLogoValue);
            var ContactNameValue=sheet.getRange(i+1,ContactNameIndex).getValue();
            var ContactNumberValue=sheet.getRange(i+1,ContactNumberIndex).getValue();
            var FooterNumberValue=sheet.getRange(i+1,FooterNumberIndex).getValue();
            
            // Create document using template
            var newSlide=DriveApp.getFileById(cloneID).makeCopy(count+' '+ClubNameValue, folderID);
            var SlideID=DriveApp.getFilesByName(count+' '+ClubNameValue).next().getId();
          
            // Append values found above to variables in slide --- update.replaceAllText('<TEXT_YOURE_SEARCHING_FOR>', <NAME_OF_THE_VALUE_VAR>);
            var update=SlidesApp.openById(SlideID);
            update.replaceAllText('{{Club Name}}', ClubNameValue);
            var image=update.getSlides()[0].getImages()[4];
            image.replace(driveImage);
            update.replaceAllText('{{Contact Name}}', ContactNameValue);
            update.replaceAllText('{{000.000.0000 x 000}}', ContactNumberValue);
            update.replaceAllText('{{000.000.0000}}', FooterNumberValue);
            
            // Up the count and alert the end
            count+=1;
        }
        ui.alert('We\'re all done!','We created '+(count-1)+' new documents for you and changed '+(count-1)*5+' variables.\nYour documents are stored on your drive in a folder called '+FolderTitle+currentTime,ui.ButtonSet.OK);
    }else{
        ui.alert('Please provide a template document ID.');
    }
}

// This Function runs once to allow for testing
function RunOne(Title, cloneID){
    Logger.log(Title+' - '+cloneID);
    if (cloneID !=  ''){
        Logger.log('Started '+cloneID);
        var TZ=CalendarApp.getDefaultCalendar().getTimeZone();
        var currentTime=Utilities.formatDate(new Date(), TZ, 'MM/dd/yyyy-HH:mm');
        Logger.log('Found Time '+currentTime);
    
        // Error checking for empty title variable
        var FolderTitle='';
        if(Title !=  ''){
            FolderTitle=Title+'-';
            Logger.log('Folder title was supplied');
        }else{
            Logger.log('No folder title was supplied');
        }
        Logger.log(FolderTitle+currentTime);
      
        //Find and define the columns --- var <ANY_NAME>=findHeader('<TITLE_OF_COLUMN>');
        var ClubNameIndex=findHeader('Club Name');
        var ClubLogoIndex=findHeader('Club Logo');
        var ContactNameIndex=findHeader('Contact Name');
        var ContactNumberIndex=findHeader('Contact Number');
        var FooterNumberIndex=findHeader('Footer Number');
        Logger.log('Passed column definition');
    
        // Create Folder using date and time found above
        DriveApp.createFolder(FolderTitle+currentTime);
        var folder=DriveApp.getFoldersByName(FolderTitle+currentTime).next().getId();
        var folderID=DriveApp.getFolderById(folder);
        folderID.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT); //Set Access to parent folder - ANYONE on challengersports.com can EDIT
        Logger.log('Passed creating the folder');
        
        // Get Values of cells to fill variables --- var <ANY_NAME>=sheet.getRange(i+1,<COLUMN_DEFINITION_VAR>).getValue();
        var i=1;
        var ClubNameValue=sheet.getRange(i+1,ClubNameIndex).getValue();
        var ClubLogoValue=sheet.getRange(i+1,ClubLogoIndex).getValue();
        var driveImage=DriveApp.getFileById(ClubLogoValue);
        var ContactNameValue=sheet.getRange(i+1,ContactNameIndex).getValue();
        var ContactNumberValue=sheet.getRange(i+1,ContactNumberIndex).getValue();
        var FooterNumberValue=sheet.getRange(i+1,FooterNumberIndex).getValue();
        Logger.log('Passed value definition');
         
        // Create document using template
        var newSlide=DriveApp.getFileById(cloneID).makeCopy(count+' '+ClubNameValue, folderID);
        var SlideID=DriveApp.getFilesByName(count+' '+ClubNameValue).next().getId();
        Logger.log('Passed document creation');
        
        // Append values found above to variables in slide --- update.replaceAllText('<TEXT_YOURE_SEARCHING_FOR>', <NAME_OF_THE_VALUE_VAR>);
        var update=SlidesApp.openById(SlideID);
        update.replaceAllText('{{Club Name}}', ClubNameValue);
        var image=update.getSlides()[0].getImages()[4];
        image.replace(driveImage);
        update.replaceAllText('{{Contact Name}}', ContactNameValue);
        update.replaceAllText('{{000.000.0000 x 000}}', ContactNumberValue);
        update.replaceAllText('{{000.000.0000}}', FooterNumberValue);
        Logger.log('Passed text appending');
        
        // Up the count and alert the end
        count+=1;
        ui.alert('We\'re all done!','We\'re all wrapped up here. The test ran successfully, and you\'re good to run the full script.',ui.ButtonSet.OK);
        Logger.log('Passed everything - It works!');
    }else{
        ui.alert('Please provide a template document ID.');
        Logger.log('Missing template ID');
    }
}

//Find the column indexes based on column header
function findHeader(Term){
    var searchString=Term;
    var ColumnIndex;
    for(var i=0; i<data[0].length; i++){
        if(data[0][i]==searchString){
            ColumnIndex=i+1;
            break;
        }
    }
    return ColumnIndex;
}