me.Hide
Unload.me
userform1.show 'userform1 is the name of the form - subject to change 

'==================================================================================================================

'********************
'    HIDE USERFORM
'********************

'in the code 

me.Hide

'Hide is the method and it will hide the entire form and all elements contained within it
'""me"" refers specifically to the useform or essentially a 'visible' module holding all the code and buttons 
'the userform is just a container for all the buttons and features, me.hide will hide it all in one fell swoop

'==================================================================================================================

'********************
'    UNLOAD USERFORM
'********************

'hide only hides the userform its still active, to get rid of it you need to use:

unload.me

'or the red ""X"" this will ""unload"" the userform

'==================================================================================================================

'***************
'    TAB ORDER
'***************

'in properties window in userform the field ""TabIndex"" is the order of areas that will receive 
'focus when you hit tab and cycle through them

'==================================================================================================================

'**************************************************
'    DEFAULT VALUES IN TEXT BOXES ON INITIALIZE
'**************************************************

'when editing userform, to have a default entry in the text box or to have something such as
'the country never change, just change the default value by entering text in the UF editor

'==================================================================================================================

'************************************
'    COLUMN WIDTHS IN COMBO BOXES
'************************************

'when doing column widths in things like combo boxes to display results, using semi-colons 
'will let you add multiple in line items of differing value

'==================================================================================================================


'************************************************************
'    HIDE APPLICATION AND ONLY USERFORM IS VISIBLE
'************************************************************

'to hide XL application and only show userform do this
'also have all closing or cancelling options show the application or a button to show application as a contingency 

sub hideWB
    application.visible = False
    userform.show
end sub

'this way the form is visible and the application isnt. make sure to build it out so application shows again

'application.visible = false also prevents adding data to a worksheet while 'false'. 
'The only workaround I found was entering data into a closed file or unhiding the application prior to data entry briefly."""
