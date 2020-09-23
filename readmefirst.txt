I wrote this Messagebox because the Messagebox provided by Microsoft VB6 did not satisfy to my requirements.

This is a graphical messagebox that you can change grapicly entirely to your wishes.

I used the properties for the buttons and skin that are used in the messageboxform, as properties of the messagebox control, so that
when you compile the code to an activeX, you don't have to show other programmers the code for the buttons,textbox and skin!!!
The Messagebox automaticly adjusts the height to the messagetext, and the width to the messagetext and buttons.
I chose to save al the calculations for the Width of the buttons and the width of the form needed for the buttons to the propertybag,
so all of the repositioning and resizing are automaticly done when you change the caption or images of the buttons and formskin at design time,
So, you don't have to take notice for the buttonwidth when the captions change, or the right positions when the formskin changes.
This also increaces the speed and avoids to much calculation at runtime.

Please do not use the graphics I made in any commercial program, I disigned these for the company that I work for, and they are therefore
protected by the strict copyrightlaws used in Belgium.
 
Please  firstly read the file “copyright.txt” before testing or using  the provided software.

Below you can find credits and more explanation.

Thanks for downloading,

Frank


 Programmer:        	Donckers Frank
                    	DarkManSoft@Gmail.com

 Description:       	Active X Control MessageBox

 Use:               	ShowMsgBox(Optional ByVal Message As String,
                   	       Optional ByVal MsgButtons As MessageButton,
                               Optional ByVal msgIcon As MessageIcons,
                               Optional Title As String,
                               Optional ByVal MSGboxType As MSGboxtypes,
                               Optional InputText As String,
                               Optional HelpButton As Boolean,
                               Optional WindowsStartUpposition As StartupPositions)
                               As String

		   	Message:
				Text to show as message

			MessageButton:
				Sets the buttons
				Values:
    					OKOnly 			= 0
    					OKCancel 		= 1
    					YesNo 			= 2
    					YesNoCancel 		= 3
    					AbortRetryIgnore 	= 4

		   	MessageIcons:
				Icon to show
				Values:
        				Questionmark 	= 1
    					Information 	= 2
    					Exclamation 	= 3
   					Critical 	= 4
               		Title:
				Sets the caption on the MessageBox
			MSGboxtypes:
				Type of Messagbox
				Values:
    					MessageBox 	= 0
    					InputBox 	= 1

			InputText: 
				Sets the inputtext when MSGboxtypes = InputBox

			HelpButton: 
				Adds a helpbutton 
				Values:
					True/False

			StartupPositions:
				Values:
					Manual = 0
    					CenterOwner 		= 1
   					CenterScreen 		= 2
	    				LeftTopOwner 		= 3
    					LeftTopScreen 		= 4
    					LeftBottomOwner 	= 5
    					LeftBottomScreen 	= 6
	    				RightTopOwner 		= 7
    					RightTopScreen 		= 8
    					RightBottomOwner	= 9
    					RightBottomScreen 	= 10
    					LastUsed 		= 11


                   Example:
                   Text1.Text = DMmsgBox1.ShowMsgBox("Message", OKOnly, , "Title", Information, MessageBox, False, LastUsed)

