var aliases = GmailApp.getAliases();
var atTest = 'ed@christchurchlondon.org';  
var globalSub = ' has expressed interest in ';
var testText = '<p>Email Recipients: ';
var nameTitle = '<b>Name:</b> '
var emailTitle = '<br/><b>Email:</b> ';
var phoneTitle =  '<br/><b>Phone:</b> ';

//Connect Groups
var cgSub = 'Connect Group Enquiry: ';
var atClapham = 'jonnyelwyn@gmail.com';
var atCommunityEngagement = 'purceada@icloud.com, faith@christchurchlondon.org, ';
var atInternationalAffairs = 'jo_nussbaum@hotmail.com, emilyatravis@googlemail.com, ';
var atSouthCentralEstates = 'stephen_humphreys@hotmail.co.uk, alicia.p.wells@googlemail.com, ';
var atSouthCentralNeighbourhoods = 'dayo65@hotmail.com, ';
var atSouthEast = 'tomhelyarcardwell@hotmail.com, vickihelyarcardwell@googlemail.com, ';
var atSouthWestNeighbourhoods = 'pip.joubert@gmail.com, werner.joubert@gmail.com>, ';
var atGreenwichAndBlackheath = 'joel.nazar@gmail.com, rebecca.gardiner@live.co.uk, ';
var atCanaryWharf = 'stella-maris@marisella.co.uk, ian@pontvert.co.uk, ';
var atGoodIdeas = 'aedwardsidun@gmail.com jg@jgosden.com, ';
var atNorthLondon = 'abcatlett@gmail.com, angiecatlett@gmail.com, ';
var atOutdoorPursuits = 'jo.wells@christchurchlondon.org, ';
var atRainford = 'dubelesly@gmail.com, ';
var atSocialJustice = 'lizzy.salway@gmail.com, sarahsandhu89@gmail.com, ';
var atSutton = 'ianrushton@hotmail.com, rushtonheather@googlemail.com, ';
var atWorkPlaces = 'hannah.r.robinson@gmail.com, ';
var atYoungAdultsCentral = 'vicki_cavolina@hotmail.com, marcasino@hotmail.com, ';
var atCentralStudents = 'jo.wells@christchurchlondon.org, ';
var atCoventGarden = 'cathyblair1@gmail.com, ';
var atHighburyAndIslington = 'donnchadhgreene@hotmail.co.uk, rorybarton@live.co.uk, ';
var atStudentsLondonBridge = 'cameron_myers@live.com, faith.kenny@yahoo.co.uk, ';
var atStudentsWestEnd = 'streetm1@live.com, ';
var atLocalCommunities = 'craigasmith11@gmail.com, naomiclarebedford@gmail.com, ';
var atCreatives = 'jack.wells@gmail.com, fionawells88@hotmail.co.uk, ';
var atNeighbourhoods = 'alastair.j.marsh@gmail.com, tatyana.marsh@gmail.com, ';
var atEmbraceEast = 'basilmussad@hotmail.co.uk, catherine.warren@hotmail.co.uk, ';
var atReachEast = 'matthew_endersby@hotmail.com, raphael_arthur@live.co.uk, ';
var atYoungAdults = 'dom_harrison@live.co.uk, deegoodfruit@gmail.com, ';
var atStudentsEast = ''; //****

//Welcome Events
var welcomeSub = ' is interested in Welcome Events at the '
var atSouthWelcome = 'katherinetait.uk@gmail.com, ';
var atCentralWelcome = 'johnpeterarcher@gmail.com, suzie.harris1@gmail.com, jo.wells@christchurchlondon.org, ';
var atWestEndWelcome = ''; //*****
var atEastWeclome = 'lozrichards1989@gmail.com, '

//Serving
var servingSub = ' is interested in serving on the ';
var atFamilies = 'amy@christchurchlondon.org, ';
var atPrayer = 'liam@christchurchlondon.org, ';
var atProduction = 'nate@christchurchlondon.org, ';
var atTech = 'nate@christchurchlondon.org, ';
var atWelcomeTeam = 'jo.wells@christchurchlondon.org, ';
var atWorship = 'rich@christchurchlondon.org, ';

//Courses
var coursesSub = 'Course Enquiry: ';
var atAlpha = '';//****
var atEmotionalHealth = 'sharon@christchurchlondon.org, ';
var atGriefShare = 'sharon@christchurchlondon.org, ';
var atMarriagePrep = ''; //****
var atMensRecovery = 'pastoralsupport@christchurchlondon.org, ';
var atPastoralSupport = 'pastoralsupport@christchurchlondon.org, ';
var atSteps = 'lars@christchurchlondon.org, ';

//Other
var atSocialAction = 'tim@christchurchlondon.org, ';
var atStudents = 'jo.wells@christchurchlondon.org, ';
var atKidsWork = 'amy@christchurchlondon.org, ';
var atNewsletter = 'ed@christchurchlondon.org';

//The Stuff
var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
var sheet = doc.getSheetByName("form-submissions");

function getFirstEmptyRow() {
  var column = sheet.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct][0] != "" ) {
    ct++;
  }
  return (ct-1);
}

var EMAIL_SENT = "EMAIL_SENT";
var xyz = getFirstEmptyRow();

function processData() {
  var sheet = doc.getSheetByName("form-submissions");
  var startRow = 2;  // First row of data to process
  var numRows = xyz;   // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 26)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    
    //Get the data
    var row = data[i];
    var title = row[1];
    var name = row[2];
    var email = row[3];
    var phone = row[4];
    var service = row[5];
    var connectgroup = row[6];
    var servingTeams = row[7]+', '+row[8]+', '+row[9]+', '+row[10]+', '+row[11]+', '+row[12];
    var courses = row[13]+', '+row[14]+', '+row[15]+', '+row[16]+', '+row[17]+', '+row[18]+', '+row[19];
    var welcomeEvents = row[20];
    var kidsWork = row[21];
    var socialAction = row[22];
    var students = row[23];
    var newsletter = row[24];
    
    var message = title + email;
    var emailSent = row[25];     // 26th column
    
    var test = 1;
    
    if (test == 1) { //Test Suite
      
      if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      
          //Thank You!
          if (email != 0) {
            GmailApp.sendEmail(email, 'Helping you Get Connected at ChristChurch London', 'Error', {noReply: true, htmlBody: '<p>Dear '+name+',</p><p>Thank you for filling in the Get Connected form and expressing interest in ChristChurch London. Members of our welcome team will be in touch with you shortly to help you find out more about us.</p><p>ChristChurch London</p>'});
          }
          
          //Connect Groups
          if (connectgroup == 'Clapham'){
            GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atClapham+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
          } else if (connectgroup == 'Community Engagement') {
            GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atCommnityEngagement+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'International Affairs') {
            GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atInternationalAffairs+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'South Central Estates') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atSouthCentralEstates+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'South Central Neighbourhoods') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atSouthCentralNeighbourhoods+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'South East') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atSouthEast+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'South West Neighbourhoods') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atSouthWestNeighbourhoods+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Greenwich and Blackheath') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atGreenwichAndBlackheath+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Canary Wharf') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atCanaryWharf+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Good Ideas') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atGoodIdeas+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'North London') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atNorthLondon+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Outdoor Pursuits') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atOutdoorPursuits+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Rainford') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atRainford+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Social Justice') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atSocialJustice+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Sutton') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atSutton+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Workplaces') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atWorkPlaces+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Young Adults Central') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atYoungAdultsCentral+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Central Students') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atCentralStudents+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Covent Garden') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atCoventGarden+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Highbury and Islington') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atHighburyAndIslington+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Students London Bridge') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atStudentsLondonBridge+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Students West End') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atStudentsWestEnd+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Local Communities') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atLocalCommunities+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Creatives') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atCreatives+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Neighbourhoods') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atNeighbourhoods+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Embrace East') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atEmbraceEast+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Reach East') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atReachEast+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Young Adults') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atYoungAdults+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (connectgroup == 'Students East') {
	        GmailApp.sendEmail(atTest, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atStudentsEast+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } 
	      
	      //Welcome events
	      if (service == 'south' /*&& welcomeEvents == 'yes-welcome-events'*/){
	        GmailApp.sendEmail(atTest, name+welcomeSub+service+' service.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atSouthWelcome+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (service == 'central' /*&& welcomeEvents == 'yes-welcome-events'*/){
	        GmailApp.sendEmail(atTest, name+welcomeSub+service+' service.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atCentralWelcome+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (service == 'west-end' /*&& welcomeEvents == 'yes-welcome-events'*/){
	        GmailApp.sendEmail(atTest, name+welcomeSub+service+' service.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atWestEndWelcome+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      } else if (service == 'east' /*&& welcomeEvents == 'yes-welcome-events'*/){
	        GmailApp.sendEmail(atTest, name+welcomeSub+service+' service.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atEastWeclome+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      }
	      
	      //Serving
	      if (row[7] == 'yes-families'){
	        GmailApp.sendEmail(atTest, name+servingSub+'Families Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atFamilies+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      if (row[8] == 'yes-prayer'){
	        GmailApp.sendEmail(atTest, name+servingSub+'Prayer Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atPrayer+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      }
	      if (row[9] == 'yes-production'){
	        GmailApp.sendEmail(atTest, name+servingSub+'Production Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atProduction+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      }
	      if (row[10] == 'yes-tech'){
	        GmailApp.sendEmail(atTest, name+servingSub+'Tech Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atTech+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      }
	      if (row[11] == 'yes-welcome'){
	        GmailApp.sendEmail(atTest, name+servingSub+'Welcome Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atWelcomeTeam+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      }
	      if (row[12] == 'yes-worship'){
	        GmailApp.sendEmail(atTest, name+servingSub+'Worship Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atWorship+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      }
	      
	      //Courses
	      if (row[13] == 'yes-alpha'){
	         GmailApp.sendEmail(atTest, coursesSub+name+globalSub+'Alpha.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atAlpha+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      if (row[14] == 'yes-emotional-health'){
	        GmailApp.sendEmail(atTest, coursesSub+name+globalSub+'Emotional Health.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atEmotionalHealth+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      if (row[15] == 'yes-grief-share'){
	        GmailApp.sendEmail(atTest, coursesSub+name+globalSub+'Grief Share.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atGriefShare+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      if (row[16] == 'yes-marriage-preparation'){
              GmailApp.sendEmail(atTest, coursesSub+name+globalSub+'Marriage Prep.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atMarriagePrep+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      if (row[17] == 'yes-mens-recovery'){
	        GmailApp.sendEmail(atTest, coursesSub+name+globalSub+'Mens Recovery.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atMensRecovery+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      if (row[18] == 'yes-pastoral-support'){
	        GmailApp.sendEmail(atTest, coursesSub+name+globalSub+'Pastoral Support.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atPastoralSupport+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
          };
          if (row[19] == 'yes-steps'){
	        GmailApp.sendEmail(atTest, coursesSub+name+globalSub+'Steps.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atSteps+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      
	      //Social Action
	      if (socialAction == 'yes-social-action'){
	        GmailApp.sendEmail(atTest, name+globalSub+' Social Action.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atSocialAction+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      //Students
	      if (students == 'yes-students'){
	        GmailApp.sendEmail(atTest, name+globalSub+' Students.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atStudents+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      //Kids
	      if (kidsWork == 'yes-kids-work'){
	        GmailApp.sendEmail(atTest, name+globalSub+' Kids Work.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atKidsWork+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
	      //Newsletter
	      if (newsletter == 'yes-newsletter'){
	        GmailApp.sendEmail(atTest, name+globalSub+' the Newsletter.', 'Error', {from: aliases[1], replyTo: email, htmlBody: '<p>'+testText+atNewsletter+'</p>'+nameTitle+name+emailTitle+email+phoneTitle+phone});
	      };
        
        sheet.getRange(startRow + i, 26).setValue(EMAIL_SENT);
        // Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
      } // end email sent
    
    } else if (test == 0) {
      
      if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
        
        //Thank you
        if (email != 0) {
          GmailApp.sendEmail(email, 'Helping you Get Connected at ChristChurch London', 'Error', {noReply: true, htmlBody: '<p>Dear '+name+',</p><p>Thank you for filling in the Get Connected form and expressing interest in ChristChurch London. Members of our welcome team will be in touch with you shortly to help you find out more about us.</p><p>ChristChurch London</p>'});
        }
        
        //Connect Groups
        if (connectgroup == 'Clapham'){
          GmailApp.sendEmail(atClapham, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Community Engagement') {
          GmailApp.sendEmail(atCommunityEngagement, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'International Affairs') {
          GmailApp.sendEmail(atInternationalAffairs, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'South Central Estates') {
          GmailApp.sendEmail(atSouthCentralEstates, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'South Central Neighbourhoods') {
          GmailApp.sendEmail(atSouthCentralNeighbourhoods, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'South East') {
          GmailApp.sendEmail(atSouthEast, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'South West Neighbourhoods') {
          GmailApp.sendEmail(atSouthWestNeighbourhoods, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Greenwich and Blackheath') {
          GmailApp.sendEmail(atGreenwichAndBlackheath, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Canary Wharf') {
          GmailApp.sendEmail(atCanaryWharf, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Good Ideas') {
          GmailApp.sendEmail(atGoodIdeas, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'North London') {
          GmailApp.sendEmail(atNorthLondon, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Outdoor Pursuits') {
          GmailApp.sendEmail(atOutdoorPursuits, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Rainford') {
          GmailApp.sendEmail(atRainford, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Social Justice') {
          GmailApp.sendEmail(atSocialJustice, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Sutton') {
          GmailApp.sendEmail(atSutton, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Workplaces') {
          GmailApp.sendEmail(atWorkPlaces, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Young Adults Central') {
          GmailApp.sendEmail(atYoungAdultsCentral, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Central Students') {
          GmailApp.sendEmail(atCentralStudents, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Covent Garden') {
          GmailApp.sendEmail(atCoventGarden, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Highbury and Islington') {
          GmailApp.sendEmail(atHighburyAndIslington, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Students London Bridge') {
          GmailApp.sendEmail(atStudentsLondonBridge, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Students West End') {
          GmailApp.sendEmail(atStudentsWestEnd, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Local Communities') {
          GmailApp.sendEmail(atLocalCommunities, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Creatives') {
          GmailApp.sendEmail(atCreatives, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Neighbourhoods') {
          GmailApp.sendEmail(atNeighbourhoods, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Embrace East') {
          GmailApp.sendEmail(atEmbraceEast, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Reach East') {
          GmailApp.sendEmail(atReachEast, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Young Adults') {
          GmailApp.sendEmail(atYoungAdults, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (connectgroup == 'Students East') {
          GmailApp.sendEmail(atStudentsEast, cgSub+name+globalSub+connectgroup, 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } 
        
        //Welcome events
        if (service == 'south' /*&& welcomeEvents == 'yes-welcome-events'*/){
          GmailApp.sendEmail(atSouthWelcome, name+welcomeSub+service+' service.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (service == 'central' /*&& welcomeEvents == 'yes-welcome-events'*/){
          GmailApp.sendEmail(atCentralWelcome, name+welcomeSub+service+' service.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (service == 'west-end' /*&& welcomeEvents == 'yes-welcome-events'*/){
          GmailApp.sendEmail(atWestEndWelcome, name+welcomeSub+service+' service.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        } else if (service == 'east' /*&& welcomeEvents == 'yes-welcome-events'*/){
          GmailApp.sendEmail(atEastWeclome, name+welcomeSub+service+' service.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        }
        
        //Serving
        if (row[7] == 'yes-families'){
          GmailApp.sendEmail(atFamilies, name+servingSub+'Families Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        if (row[8] == 'yes-prayer'){
          GmailApp.sendEmail(atPrayer, name+servingSub+'Prayer Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        if (row[9] == 'yes-production'){
          GmailApp.sendEmail(atProduction, name+servingSub+'Production Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        if (row[10] == 'yes-tech'){
          GmailApp.sendEmail(atTech, name+servingSub+'Tech Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        if (row[11] == 'yes-welcome'){
          GmailApp.sendEmail(atWelcomeTeam, name+servingSub+'Welcome Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        if (row[12] == 'yes-worship'){
          GmailApp.sendEmail(atWorship, name+servingSub+'Worship Team.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        
        //Courses
        if (row[13] == 'yes-alpha'){
          GmailApp.sendEmail(atAlpha, coursesSub+name+globalSub+'Alpha.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        if (row[14] == 'yes-emotional-health'){
          GmailApp.sendEmail(atEmotionalHealth, coursesSub+name+globalSub+'Emotional Health.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        if (row[15] == 'yes-grief-share'){
          GmailApp.sendEmail(atGriefShare, coursesSub+name+globalSub+'Grief Share.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        if (row[16] == 'yes-marriage-preparation'){
          GmailApp.sendEmail(atMarriagePrep, coursesSub+name+globalSub+'Marriage Prep.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
	    };
        if (row[17] == 'yes-mens-recovery'){
          GmailApp.sendEmail(atMensRecovery, coursesSub+name+globalSub+'Mens Recovery.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
	    };
        if (row[18] == 'yes-pastoral-support'){
          GmailApp.sendEmail(atPastoralSupport, coursesSub+name+globalSub+'Pastoral Support.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        if (row[19] == 'yes-steps'){
          GmailApp.sendEmail(atSteps, coursesSub+name+globalSub+'Steps.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        
        //Social Action
        if (socialAction == 'yes-social-action'){
          GmailApp.sendEmail(atSocialAction, name+globalSub+' Social Action.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        //Students
        if (students == 'yes-students'){
          GmailApp.sendEmail(atStudents, name+globalSub+' Students.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        //Kids
        if (kidsWork == 'yes-kids-work'){
          GmailApp.sendEmail(atKidsWork, name+globalSub+' Kids Work.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});
        };
        //Kids
        if (newsletter == 'yes-newsletter'){
          GmailApp.sendEmail(atNewsletter, name+globalSub+' the Newsletter.', 'Error', {from: aliases[1], replyTo: email, htmlBody: nameTitle+name+emailTitle+email+phoneTitle+phone});          
        };
        
        sheet.getRange(startRow + i, 26).setValue(EMAIL_SENT);
        // Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
      }//end email sent
    
   }//end test if
  
  }//end row iterate
}//end function