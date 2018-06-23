
function onOpen2(){  

  var ui = SpreadsheetApp.getUi();
	ui.createMenu("Index")
    .addItem("1 Call Events", "pEv")
    .addItem("2 Write Event ID", "eventQA2")
	.addItem("3 Plan Copywriters", "pCopy")
    .addItem("4 Push to Copywriter", "qCopy")
    .addItem("5 Populate the Hopper", "uHop2")
    .addItem("6 Process the Hopper", "uHop")
    .addItem("7 Clear Sendwithus", "SWUclear")
    .addItem("8 Write to Sendwithus", "Email")
    .addItem("9 User Data", "data")
	.addToUi();
  
  
}

function pEv(){
 
  var iEventID = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID"); // locates Event ID sheet
  var ePs2 = iEventID.getRange("ad3:ad1991").getValues(); // creates an array of column 30 in Event ID sheet where the holding ID of events that have yet to be QA'd
  var ePfull = iEventID.getRange("a3:as1991").getValues(); // takes all Event ID and makes it an array
  var evQueue = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event QA"); // Locates the Event QA sheet
  
  for (var i=2; i < 1990; i++){
    
    if(ePs2[i] > 0){
     
      var print2 = evQueue.getRange(2,1).getValue(); // identifies where to print the details
      var xContact = ePfull[i][5]; // contact
      var xAccount = ePfull[i][6]; // account
      var xEvent = ePfull[i][13]; // event note
      var xLabel = ePfull[i][14]; // event label
      var xUEmail = ePfull[i][10]; // user email
      var xEURL = ePfull[i][15]; // event URL
      var xID = ePfull[i][0]; // event ID - maybe also i+1
      var xUser = ePfull[i][1]; // user 
      var xCompany = ePfull[i][2]; // user's compay
      var xSheet = ePfull[i][3]; // user account sheet link
      var xLink = ePfull[i][7]; // contact LinkedIn
      var xCEmail = ePfull[i][8]; // contact email 
      var xDate = ePfull[i][16]; // event Date
      var xRole = ePfull[i][12]; // contact Role
      
      evQueue.getRange(print2,1).setValue(xID);
      evQueue.getRange(print2,2).setValue(xUser);
      evQueue.getRange(print2,3).setValue(xCompany);
      evQueue.getRange(print2,4).setValue(xContact);
      evQueue.getRange(print2,5).setValue(xAccount);
      evQueue.getRange(print2,6).setValue(xLink);
      evQueue.getRange(print2,7).setValue(xRole);
      evQueue.getRange(print2,8).setValue(xEvent);
      evQueue.getRange(print2,9).setValue(xLabel);
      evQueue.getRange(print2,10).setValue(xEURL);
      evQueue.getRange(print2,11).setValue(xDate);
      
    }
    
  }   
  
}
  
function pCopy(){
  
  var iEventID = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID"); // locates Event ID sheet
  var ePs = iEventID.getRange("aj3:aj1991").getValues(); // creates an array of column 36 in Event ID sheet where the holding ID of to be drafted events is kept
  var ePfull = iEventID.getRange("a3:bj1991").getValues();
  
  for (var i=2; i < 1990; i++){
    
    if(ePs[i] > 0){
     
      var pEvents = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Passed Events"); // Locates the Passed Events sheet
      var print = pEvents.getRange(1,1).getValue(); // identifies where to print the details
      var xUser = ePfull[i][1];
      var xCompany = ePfull[i][2];
      var xEvent = ePfull[i][13];
      var xLabel = ePfull[i][14];
      pEvents.getRange(print+2,2).setValue(i+1);
      pEvents.getRange(print+2,4).setValue(xUser);
      pEvents.getRange(print+2,5).setValue(xCompany);
      pEvents.getRange(print+2,6).setValue(xEvent);
      pEvents.getRange(print+2,7).setValue(xLabel);
     
      
    }
    
  } 

}





function qCopy() {
  
  var Karen = SpreadsheetApp.openById("11Vy-fUMSHGbkXf5hnlIHna1uCNHhK7uPT4rPSwxZf6Y").getSheetByName("Event Queue");
  var Ronnie = SpreadsheetApp.openById("1ixgXOMUT9PHOyKeV5K1XCEmGnSg-B2k0hjQtQyZn3FU").getSheetByName("Event Queue");
  var Alison = SpreadsheetApp.openById("1O1t3I_BYILVjmiLXPeu_BGWIYU009XixjsdD1A8DOVM").getSheetByName("Event Queue");
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var passEvent = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Passed Events");
  var cWriters = passEvent.getRange("b2:C102").getValues(); // copywriter IDs
  var ePfull = eventIDss.getRange("a3:bj1991").getValues(); // takes all Event ID and makes it an array
  
  for(var i = 0; i < 100; i++){
    
      var writer = cWriters[i][1];
    
    
    if(writer == "Ronnie"){
      
      var preID = cWriters[i][0]; // finds the ID of the event that has been allotted to this writer
      var eID = preID-1;
      Logger.log(eID);
      var xContact = ePfull[eID][5]; // contact
      var xAccount = ePfull[eID][6]; // account
      var xEvent = ePfull[eID][13]; // event note
      var xLabel = ePfull[eID][14]; // event label
      var xUEmail = ePfull[eID][10]; // user email
      var xEURL = ePfull[eID][15]; // event URL
      var xID = ePfull[eID][0]; // event ID - maybe also i+1
      var xUser = ePfull[eID][1]; // user 
      var xCompany = ePfull[eID][2]; // user's compay
      var xSheet = ePfull[eID][3]; // user account sheet link
      var xLink = ePfull[eID][7]; // contact LinkedIn
      var xCEmail = ePfull[eID][8]; // contact email 
      var xDate = ePfull[eID][16]; // event Date
      var xRole = ePfull[eID][12]; // contact Role
      //var xRole = ePfull[eID][12]; // account homepage
       
      
      var ronps = Ronnie.getRange(1,1).getValue();
      var ronp = ronps*10;
      
      Ronnie.getRange(ronp+2,1).setValue(xID);
      Ronnie.getRange(ronp+2,2).setValue(xUser);
      Ronnie.getRange(ronp+2,3).setValue(xCompany);
      Ronnie.getRange(ronp+2,4).setValue(xContact);
      Ronnie.getRange(ronp+2,5).setValue(xAccount);
      Ronnie.getRange(ronp+2,6).setValue(xDate);
      Ronnie.getRange(ronp+4,2).setValue(xEvent);
      Ronnie.getRange(ronp+4,4).setValue(xEURL);
      Ronnie.getRange(ronp+4,6).setValue(xLink);
      Ronnie.getRange(ronp+4,5).setValue(xLabel);
     
      
      eventIDss.getRange(preID+2,37).setValue("Ronnie");
     
      
    } else if(writer == "Karen"){
        
      Logger.log(writer);
      var preID = cWriters[i][0]; // finds the ID of the event that has been allotted to this writer
      var eID = preID-1;
      Logger.log(eID);
      var xContact = ePfull[eID][5]; // contact
      var xAccount = ePfull[eID][6]; // account
      var xEvent = ePfull[eID][13]; // event note
      var xLabel = ePfull[eID][14]; // event label
      var xUEmail = ePfull[eID][10]; // user email
      var xEURL = ePfull[eID][15]; // event URL
      var xID = ePfull[eID][0]; // event ID - maybe also i+1
      var xUser = ePfull[eID][1]; // user 
      var xCompany = ePfull[eID][2]; // user's compay
      var xSheet = ePfull[eID][3]; // user account sheet link
      var xLink = ePfull[eID][7]; // contact LinkedIn
      var xCEmail = ePfull[eID][8]; // contact email 
      var xDate = ePfull[eID][16]; // event Date
      var xRole = ePfull[eID][12]; // contact Role
      //var xRole = ePfull[eID][12]; // account homepage
      
      
      var karens = Karen.getRange(1,1).getValue();
      var karp = karens*10;
      
      Karen.getRange(karp+2,1).setValue(xID);
      Karen.getRange(karp+2,2).setValue(xUser);
      Karen.getRange(karp+2,3).setValue(xCompany);
      Karen.getRange(karp+2,4).setValue(xContact);
      Karen.getRange(karp+2,5).setValue(xAccount);
      Karen.getRange(karp+2,6).setValue(xDate);
      Karen.getRange(karp+4,2).setValue(xEvent);
      Karen.getRange(karp+4,4).setValue(xEURL);
      Karen.getRange(karp+4,6).setValue(xLink);
      Karen.getRange(karp+4,5).setValue(xLabel);
     
      
      eventIDss.getRange(preID+2,37).setValue("Karen");  
      
    }  else if(writer == "Alison"){
        
      Logger.log(writer);
      var preID = cWriters[i][0]; // finds the ID of the event that has been allotted to this writer
      var eID = preID-1;
      Logger.log(eID);
      var xContact = ePfull[eID][5]; // contact
      var xAccount = ePfull[eID][6]; // account
      var xEvent = ePfull[eID][13]; // event note
      var xLabel = ePfull[eID][14]; // event label
      var xUEmail = ePfull[eID][10]; // user email
      var xEURL = ePfull[eID][15]; // event URL
      var xID = ePfull[eID][0]; // event ID - maybe also i+1
      var xUser = ePfull[eID][1]; // user 
      var xCompany = ePfull[eID][2]; // user's compay
      var xSheet = ePfull[eID][3]; // user account sheet link
      var xLink = ePfull[eID][7]; // contact LinkedIn
      var xCEmail = ePfull[eID][8]; // contact email 
      var xDate = ePfull[eID][16]; // event Date
      var xRole = ePfull[eID][12]; // contact Role
      //var xRole = ePfull[eID][12]; // account homepage
      
      
      var alisons = Alison.getRange(1,1).getValue();
      var ali = alisons*10;
      
      Alison.getRange(ali+2,1).setValue(xID);
      Alison.getRange(ali+2,2).setValue(xUser);
      Alison.getRange(ali+2,3).setValue(xCompany);
      Alison.getRange(ali+2,4).setValue(xContact);
      Alison.getRange(ali+2,5).setValue(xAccount);
      Alison.getRange(ali+2,6).setValue(xDate);
      Alison.getRange(ali+4,2).setValue(xEvent);
      Alison.getRange(ali+4,4).setValue(xEURL);
      Alison.getRange(ali+4,6).setValue(xLink);
      Alison.getRange(ali+4,5).setValue(xLabel);
     
      
      eventIDss.getRange(preID+2,37).setValue("Alison");  
      
    }  
  
  
 }
  
  passEvent.getRange("b2:j250").clearContent();
  
}


function eventQA2() {
  
  var eventQAss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event QA");
  var eventIDss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event ID");
  var qArea = eventQAss.getRange("a3:m253").getValues();
  
  Logger.log(qArea[2][4]);
  
  for(var i = 0; i < 250; i++){
    
    if(qArea[i][12] !== ""){    // this is the QA label column
      
      Logger.log(qArea[i][12]);
      Logger.log(qArea[i][0]);
      
     eventIDss.getRange(qArea[i][0]+2,27).setValue(qArea[i][12]); // writes the event QA label to Event ID
     eventIDss.getRange(qArea[i][0]+2,26).setValue(qArea[i][11]); // writes the event QA notes to Event ID
     eventIDss.getRange(qArea[i][0]+2,14).setValue(qArea[i][7]); // writes the event QA notes to Event ID 
      
     var eM = eventIDss.getRange(qArea[i][0]+2,18).getValue();
     var qA = eventIDss.getRange(qArea[i][0]+2,27).getValue();
     var qAComm = eventIDss.getRange(qArea[i][0]+2,26).getValue();
     var desc = eventIDss.getRange(qArea[i][0]+2,31).getValue();
     var user = eventIDss.getRange(qArea[i][0]+2,2).getValue();
     var id = eventIDss.getRange(qArea[i][0]+2,1).getValue();
      
     MailApp.sendEmail(eM,user, desc + "      " + qA + "     " + qAComm + "    Event ID " + id);
     
    }
   
 
  }
  
 eventQAss.getRange("A3:m253").clearContent();
  
}


function uHop(){
  
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var hopper = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Hopper");
  var ePfull = eventIDss.getRange("a3:bj1991").getValues(); // takes all Event ID and makes it an array
  var qcSearch = hopper.getRange("a2:c400").getValues();
  var swu = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Sendwithus");
  
  for (var b = 0; b < 390; b++){
   
    var hop = qcSearch[b][2];
    
    if(hop =="Push to Email"){
     
      var nHead = hopper.getRange(b-2,5).getValue();
      var nGreet = hopper.getRange(b-1,5).getValue();
      var nBody = hopper.getRange(b,5).getValue();
      var eNum = hopper.getRange(b-5,1).getValue();
      var note = hopper.getRange(b+1,4).getValue();
      var ch1 = hopper.getRange(b-2,4).getValue();
      var ch2 = hopper.getRange(b-1,4).getValue();
      var ch3 = hopper.getRange(b,4).getValue();
      var eID = ePfull[eNum-1][0];
      var cEm = ePfull[eNum-1][40];
        
      if(ch1 !== "" || ch2 !== "" || ch3 !== ""){    // if any of the overrides (edit boxes) have been used then var ch is declared as "Change" 
                                                     // note logical operators || is OK, && is AND, ! is NOT
       var ch = "Change";
        
      MailApp.sendEmail(cEm + "," + "denis.anscomb@gmail.com", "Change made to bmail to " 
                        + ePfull[eNum-1][1] + " for event " + eNum, "Hi " + ePfull[eNum-1][36] 
                        + "  The copy linked to this event was edited:  " 
                        + ePfull[eNum-1][30] + "  " + "                 From:                              "+ ePfull[eNum-1][32] + "  "
                        + ePfull[eNum-1][33] + "  " + ePfull[eNum-1][34]
                        + "         To:                      " + nHead + " . " +nGreet + "  " + nBody 
                        + "            Note:                            " + note);  // feedback email to writers
      
      }
      
      var qc = qcSearch[b-7][0];
      Logger.log(qc);
      Logger.log(b); 
      
      eventIDss.getRange(qc+2,49).setValue(nHead);
      eventIDss.getRange(qc+2,53).setValue(nGreet);
      eventIDss.getRange(qc+2,54).setValue(nBody);
      eventIDss.getRange(qc+2,38).setValue("email");
      eventIDss.getRange(qc+2,56).setValue(note);
      
      hopper.getRange(b-2,4).clearContent();
      hopper.getRange(b-1,4).clearContent();
      hopper.getRange(b,4).clearContent();
      hopper.getRange(b-2,3).clearContent();
      hopper.getRange(b-1,3).clearContent();
      hopper.getRange(b,3).clearContent();
      hopper.getRange(b-5,1).clearContent();
      hopper.getRange(b-5,2).clearContent();
      hopper.getRange(b-5,3).clearContent();
      hopper.getRange(b-5,4).clearContent();
      hopper.getRange(b-5,5).clearContent();
      hopper.getRange(b-5,6).clearContent();
      hopper.getRange(b-3,2).clearContent();
      hopper.getRange(b-3,4).clearContent();
      hopper.getRange(b-3,5).clearContent();
      hopper.getRange(b-3,6).clearContent();
      hopper.getRange(b+2,5).clearContent();
      hopper.getRange(b+2,3).clearContent();
      hopper.getRange(b+1,4).clearContent();
      

      
    }else if(hop =="Leave in Queue"){
      
      var nHead = hopper.getRange(b-2,5).getValue();
      var nGreet = hopper.getRange(b-1,5).getValue();
      var nBody = hopper.getRange(b,5).getValue();
      var note = hopper.getRange(b+1,4).getValue();
      
      var qc = qcSearch[b-7][0];
      Logger.log(qc);
      Logger.log(b);
      
      eventIDss.getRange(qc+2,49).setValue(nHead);
      eventIDss.getRange(qc+2,53).setValue(nGreet);
      eventIDss.getRange(qc+2,54).setValue(nBody);
      eventIDss.getRange(qc+2,56).setValue(note);
      
      hopper.getRange(b-2,4).clearContent();
      hopper.getRange(b-1,4).clearContent();
      hopper.getRange(b,4).clearContent();
      hopper.getRange(b-2,3).clearContent();
      hopper.getRange(b-1,3).clearContent();
      hopper.getRange(b,3).clearContent();
      hopper.getRange(b-5,1).clearContent();
      hopper.getRange(b-5,2).clearContent();
      hopper.getRange(b-5,3).clearContent();
      hopper.getRange(b-5,4).clearContent();
      hopper.getRange(b-5,5).clearContent();
      hopper.getRange(b-5,6).clearContent();
      hopper.getRange(b-3,2).clearContent();
      hopper.getRange(b-3,4).clearContent();
      hopper.getRange(b-3,5).clearContent();
      hopper.getRange(b-3,6).clearContent();
      hopper.getRange(b+2,5).clearContent();
      hopper.getRange(b+2,3).clearContent();
      hopper.getRange(b+1,4).clearContent();
      
      
    } else if(hop == "Archive"){
      
      var nHead = hopper.getRange(b-2,5).getValue();
      var nGreet = hopper.getRange(b-1,5).getValue();
      var nBody = hopper.getRange(b,5).getValue();
      var note = hopper.getRange(b+1,4).getValue();
      
      var qc = qcSearch[b-7][0];
      Logger.log(qc);
      Logger.log(b); 
      
      eventIDss.getRange(qc+2,49).setValue(nHead);
      eventIDss.getRange(qc+2,53).setValue(nGreet);
      eventIDss.getRange(qc+2,54).setValue(nBody);
      eventIDss.getRange(qc+2,52).setValue("Archive");
      eventIDss.getRange(qc+2,32).setValue("");
      eventIDss.getRange(qc+2,38).setValue("");
      eventIDss.getRange(qc+2,56).setValue(note);
      
      hopper.getRange(b-2,4).clearContent();
      hopper.getRange(b-1,4).clearContent();
      hopper.getRange(b,4).clearContent();
      hopper.getRange(b-2,3).clearContent();
      hopper.getRange(b-1,3).clearContent();
      hopper.getRange(b,3).clearContent();
      hopper.getRange(b-5,1).clearContent();
      hopper.getRange(b-5,2).clearContent();
      hopper.getRange(b-5,3).clearContent();
      hopper.getRange(b-5,4).clearContent();
      hopper.getRange(b-5,5).clearContent();
      hopper.getRange(b-5,6).clearContent();
      hopper.getRange(b-3,2).clearContent();
      hopper.getRange(b-3,4).clearContent();
      hopper.getRange(b-3,5).clearContent();
      hopper.getRange(b-3,6).clearContent();
      hopper.getRange(b+2,5).clearContent();
      hopper.getRange(b+2,3).clearContent();
      hopper.getRange(b+1,4).clearContent();
    
  }
  }
}
  
  
 function uHop2(){ 
   
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var hopper = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Hopper");
  var ePfull = eventIDss.getRange("a3:bf1991").getValues(); // takes all Event ID and makes it an array
  var qcSearch = hopper.getRange("a2:c400").getValues();
  var swu = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Sendwithus");
  
  
  for(var i = 0; i < 1990; i++){
    
      var drCop = ePfull[i][31]; // looking for any Events in drafted status. 
      var eIDpre = ePfull[i][0];
      var eID = eIDpre-1;
    
    if(drCop == "Drafted"){
      
      Logger.log(i);
      
      
      var xContact = ePfull[eID][5]; // contact
      var xAccount = ePfull[eID][6]; // account
      var xEvent = ePfull[eID][13]; // event note
      var xLabel = ePfull[eID][14]; // event label
      var xUEmail = ePfull[eID][10]; // user email
      var xEURL = ePfull[eID][15]; // event URL
      var xID = ePfull[eID][0]; // event ID - maybe also i+1
      var xUser = ePfull[eID][1]; // user 
      var xCompany = ePfull[eID][2]; // user's compay
      var xSheet = ePfull[eID][3]; // user account sheet link
      var xLink = ePfull[eID][7]; // contact LinkedIn
      var xCEmail = ePfull[eID][8]; // contact email 
      var xDate = ePfull[eID][16]; // event Date
      var xRole = ePfull[eID][12]; // contact Role
      var xHead = ePfull[eID][32]; // Head
      var xGreet = ePfull[eID][33]; // greet
      var xBody = ePfull[eID][34]; // body
      var xFHead = ePfull[eID][48]; // Final Header
      var xFBody = ePfull[eID][50]; // Final Body
      var xNote = ePfull[eID][55]; // QA Note
      
      
      var hPpre = hopper.getRange(1,1).getValue();
      var hp = hPpre*10;
      
      Logger.log(hp);
      
      hopper.getRange(hp+2,1).setValue(xID);
      hopper.getRange(hp+2,2).setValue(xUser);
      hopper.getRange(hp+2,3).setValue(xCompany);
      hopper.getRange(hp+2,4).setValue(xContact);
      hopper.getRange(hp+2,5).setValue(xAccount);
      hopper.getRange(hp+2,6).setValue(xDate);
      hopper.getRange(hp+4,2).setValue(xEvent);
      hopper.getRange(hp+4,4).setValue(xEURL);
      hopper.getRange(hp+4,6).setValue(xLink);
      hopper.getRange(hp+4,5).setValue(xLabel);
      hopper.getRange(hp+5,3).setValue(xHead);
      hopper.getRange(hp+9,5).setValue(xRole);
      hopper.getRange(hp+6,3).setValue(xGreet);
      hopper.getRange(hp+7,3).setValue(xBody);
     // hopper.getRange(hp+5,4).setValue(xFHead);
     // hopper.getRange(hp+7,4).setValue(xFBody);
      hopper.getRange(hp+8,4).setValue(xNote);
      

    }
  }
  
  }

function Email(){
  
  
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var ePfull = eventIDss.getRange("a3:bf1990").getValues(); // takes all Event ID and makes it an array
  var swu = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Sendwithus");
  var renew = swu.getRange("a3:a400");
  renew.clear();
  
  
  for(var m = 0; m < 1990; m++){
    
      var yahoo = ePfull[m][37];
      var eIDpre = ePfull[m][0];
      
    
    if(yahoo == "email"){
      
      
      var checker = ePfull[m][34];
      
      
      var point = swu.getRange(2,1).getValue();
      var stick = (point*24)+3;
      var n = m;
      
      var xContact = ePfull[n][5]; // contact
      var xAccount = ePfull[n][6]; // account
      var xEvent = ePfull[n][13]; // event note
      var xLabel = ePfull[n][14]; // event label
      var xUEmail = ePfull[n][10]; // user email
      var xEURL = ePfull[n][15]; // event URL
      var xID = ePfull[n][0]; // event ID - maybe also i+1
      var xUser = ePfull[n][1]; // user 
      var xCompany = ePfull[n][2]; // user's compay
      var xSheet = ePfull[n][3]; // user account sheet link
      var xLink = ePfull[n][7]; // contact LinkedIn
      var xCEmail = ePfull[n][8]; // contact email 
      var xDate = ePfull[n][16]; // event Date
      var xRole = ePfull[n][12]; // contact Role
      var xHead = ePfull[n][48]; // Head
      var xGreet = ePfull[n][49]; // greet
      var xBody = ePfull[n][50]; // body
      var xDesc = ePfull[n][30]; // description for news
      
      Logger.log(xHead);
      Logger.log(xBody);
      Logger.log(xGreet);
      
       swu.getRange(stick,3).setValue(xUser);
       swu.getRange(stick,1).setValue(xID);
       swu.getRange(stick+1,6).setValue(xUEmail);
       swu.getRange(stick+2,3).setValue(xDesc);
       swu.getRange(stick+3,3).setValue(xCEmail);
       swu.getRange(stick+4,4).setValue(xCompany);
       swu.getRange(stick+5,4).setValue(xContact);
       swu.getRange(stick+9,3).setValue(xHead);
       swu.getRange(stick+11,3).setValue(xBody);
  
}

  }
}

function SWUclear() {
  
 var cSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sendwithus");
  
  for (var i = 0; i < 40; i++){
    
   var j = i*24 
   cSS.getRange(j+3,1).clear();
   cSS.getRange(j+3,3).clear();
   cSS.getRange(j+4,6).clear();
   cSS.getRange(j+5,3).clear();
   cSS.getRange(j+6,3).clear();
   cSS.getRange(j+7,4).clear();
   cSS.getRange(j+8,4).clear();
   cSS.getRange(j+12,3).clear();
   cSS.getRange(j+14,3).clear();
    
  }
  
}

function data() {
  
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var ePfull = eventIDss.getRange("a3:bf1991").getValues(); // takes all Event ID and makes it an array
  var users = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User List"); 
  var user = users.getRange("j2:j36").getValues(); // makes an array of the user list
  
  for (var f = 0; f < 34; f++){
    
    if(user[f][0] !==""){
      
      var cust = user[f][0]; // sets cust to be the user name for each of the users
      var count = 0;
      var cPass = 0;
  
   for(var d = 0; d < 1980; d++){
     
     if(ePfull[d][1] == cust){
          
      var count = count + 1; 
      var eventQAL = ePfull[d][29]; // should be the QA label but check
      var ev1 = ePfull[d][28]; // should be the QA label but check 
   
       if (ev1 == "PASS"){
         var cPass = cPass + 1;}
       
       users.getRange(f+2,11).setValue(count);
       users.getRange(f+2,12).setValue(cPass);
  
      
       }
      } 
     }
   }
  }



