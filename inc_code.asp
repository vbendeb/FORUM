<%
'#################################################################################
'##  Replaced by inc_code.js
'##  Leaving for use by MODS that may still be using it.
'#################################################################################
'##
'#################################################################################
'## Copyright (C) 2000-02 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or any later version.
'##
'## All copyright notices regarding Snitz Forums 2000
'## must remain intact in the scripts and in the outputted HTML
'## The "powered by" text/logo with a link back to
'## http://forum.snitz.com in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'##
'## Support can be obtained from support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## reinhold@bigfoot.com
'##
'## or
'##
'## Snitz Communications
'## C/O: Michael Anderson
'## PO Box 200
'## Harpswell, ME 04079
'#################################################################################
%>
<script language="JavaScript" type="text/javascript">

helpstat = false;
stprompt = false;
basic = true;


function thelp(swtch){
	if (swtch == 1){
		basic = false;
		stprompt = false;
		helpstat = true;
	} else if (swtch == 0) {
		helpstat = false;
		stprompt = false;
		basic = true;
	} else if (swtch == 2) {
		helpstat = false;
		basic = false;
		stprompt = true;
	}
}

function AddText(NewCode) {
document.PostTopic.Message.value+=NewCode
}

function email() {
	if (helpstat) {
		alert("E-mail Tag Turns an e-mail address into a mailto hyperlink.\n\nUSE #1: [url]someone\@anywhere.com[/url] \nUSE #2: [url=\"someone\@anywhere.com\"]link text[/url]");
		}
	else { 
		txt2=prompt("Text to be shown for the link. Leave blank if you want the e-mail address to be shown for the link.",""); 
		if (txt2!=null) {
			txt=prompt("URL for the link.","mailto:");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt=txt;
					AddText(AddTxt);
				} else {
					AddTxt="[url=\""+txt+"\"]"+txt2+"[/url]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}
function showsize(size) {
	if (helpstat) {
		alert("Size Tag Sets the text size. Possible values are 1 to 6.\n1 being the smallest and 6 the largest.\n\nUSE: [size="+size+"]This is size "+size+" text[/size="+size+"]");
	} else if (basic) {
		AddTxt="[size="+size+"][/size="+size+"]";
		AddText(AddTxt);
	} else {                       
		txt=prompt("Text to be size "+size,"Text"); 
		if (txt!=null) {             
			AddTxt="[size="+size+"]"+txt+"[/size="+size+"]";
			AddText(AddTxt);
		}        
	}
}

function bold() {
	if (helpstat) {
		alert("Bold Tag Makes the enlosed text bold.\n\nUSE: [b]This is some bold text[/b]");
	} else if (basic) {
		AddTxt="[b][/b]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made BOLD.","Text");     
		if (txt!=null) {           
			AddTxt="[b]"+txt+"[/b]";
			AddText(AddTxt);
		}       
	}
}

function italicize() {
	if (helpstat) {
		alert("Italicize Tag Makes the enlosed text italicized.\n\nUSE: [i]This is some italicized text[/i]");
	} else if (basic) {
		AddTxt="[i][/i]";
		AddText(AddTxt);
	} else {   
		txt=prompt("Text to be italicized","Text");     
		if (txt!=null) {           
			AddTxt="[i]"+txt+"[/i]";
			AddText(AddTxt);
		}	        
	}
}

function quote() {
	if (helpstat){
		alert("Quote tag Quotes the enclosed text to reference something specific that someone has posted.\n\nUSE: [quote]This is a quote[/quote]");
	} else if (basic) {
		AddTxt=" [quote] [/quote]";
		AddText(AddTxt);
	} else {   
		txt=prompt("Text to be quoted","Text");     
		if(txt!=null) {          
			AddTxt=" [quote] "+txt+" [/quote]";
			AddText(AddTxt);
		}	        
	}
}

function showcolor(color) {
	if (helpstat) {
		alert("Color Tag Sets the text color. Any named color can be used.\n\nUSE: ["+color+"]This is some "+color+" text[/"+color+"]");
	} else if (basic) {
		AddTxt="["+color+"][/"+color+"]";
		AddText(AddTxt);
	} else {  
     	txt=prompt("Text to be "+color,"Text");
		if(txt!=null) {
			AddTxt="["+color+"]"+txt+"[/"+color+"]";
			AddText(AddTxt);        
		} 
	}
}

function center() {
 	if (helpstat) {
		alert("Centered tag Centers the enclosed text.\n\nUSE: [center]This text is centered[/center]");
	} else if (basic) {
		AddTxt="[center][/center]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be centered","Text");     
		if (txt!=null) {          
			AddTxt="[center]"+txt+"[/center]";
			AddText(AddTxt);
		}	       
	}
}

function hyperlink() {
	if (helpstat) {
		alert("Hyperlink Tag \nTurns an url into a hyperlink.\n\nUSE: [url]http://www.anywhere.com[/url]\n\nUSE: [url=http://www.anywhere.com]link text[/url]");
	} else if (basic) {
		AddTxt="[url][/url]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("Text to be shown for the link.\nLeave blank if you want the url to be shown for the link.",""); 
		if (txt2!=null) {
			txt=prompt("URL for the link.","http://");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[url]"+txt+"[/url]";
					AddText(AddTxt);
				} else {
					AddTxt="[url=\""+txt+"\"]"+txt2+"[/url]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function image() {
	if (helpstat){
		alert("Image Tag Inserts an image into the post.\n\nUSE: [img]http://www.anywhere.com/image.gif[/img]");
	} else if (basic) {
		AddTxt="[img][/img]";
		AddText(AddTxt);
	} else {  
		txt=prompt("URL for graphic","http://");    
		if(txt!=null) {            
			AddTxt="[img]"+txt+"[/img]";
			AddText(AddTxt);
		}	
	}
}

function showcode() {
	if (helpstat) {
		alert("Code Tag Blockquotes the text you reference and preserves the formatting.\nUsefull for posting code.\n\nUSE: [code]This is formated text[/code]");
	} else if (basic) {
		AddTxt=" [code] [/code]";
		AddText(AddTxt);
	} else {   
		txt=prompt("Enter code","");     
		if (txt!=null) {          
			AddTxt="[code]"+txt+"[/code]";
			AddText(AddTxt);
		}	       
	}
}

function list() {
	if (helpstat) {
		alert("List Tag Builds a bulleted, numbered, or alphabetical list.\n\nUSE: [list] [*]item1[/*] [*]item2[/*] [*]item3[/*] [/list]");
	} else if (basic) {
		AddTxt=" [list][*]  [/*][*]  [/*][*]  [/*][/list]";
		AddText(AddTxt);
	} else {  
		type=prompt("Type of list Enter \'A\' for alphabetical, \'1\' for numbered, Leave blank for bulleted.","");               
		while ((type!="") && (type!="A") && (type!="a") && (type!="1") && (type!=null)) {
			type=prompt("ERROR! The only possible values for type of list are blank 'A' and '1'.","");               
		}
		if (type!=null) {
			if (type=="") {
				AddTxt="[list]";
			} else {
				AddTxt="[list="+type+"]";
			} 
			txt="1";
			while ((txt!="") && (txt!=null)) {
				txt=prompt("List item Leave blank to end list",""); 
				if (txt!="") {             
					AddTxt+="[*]"+txt+"[/*]"; 
				}                   
			} 
			if (type=="") {
				AddTxt+="[/list] ";
			} else {
				AddTxt+="[/list="+type+"]";
			} 
			AddText(AddTxt); 
		}
	}
}

function underline() {
  	if (helpstat) {
		alert("Underline Tag Underlines the enclosed text.\n\nUSE: [u]This text is underlined[/u]");
	} else if (basic) {
		AddTxt="[u][/u]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be Underlined.","Text");     
		if (txt!=null) {           
			AddTxt="[u]"+txt+"[/u]";
			AddText(AddTxt);
		}	        
	}
}

function showfont(font) {
 	if (helpstat){
		alert("Font Tag Sets the font face for the enclosed text.\n\nUSE: [font="+font+"]The font of this text is "+font+"[/font]");
	} else if (basic) {
		AddTxt="[font="+font+"][/font="+font+"]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("Text to be in "+font,"Text");
		if (txt!=null) {             
			AddTxt="[font="+font+"]"+txt+"[/font="+font+"]";
			AddText(AddTxt);
		}        
	}  
}
</SCRIPT>