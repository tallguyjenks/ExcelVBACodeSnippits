"'=====================
'VBA CLASS MODULES
'=====================

'https://youtu.be/MjbmsVDnAL0?list=PLNIs-AWhQzckr8Dgmgb3akx_gFMnpxTN5
'-----------------------------------
'What is a class
'-----------------------------------
'a blueprint of a objects, functions variables 
'a way to hide a lot of complex code behind the scenes to reference in normal modules


'-----------------------------------
'why use class modules
'-----------------------------------
'allows you to develop your own blueprints about how objects should work and what other
'users can do with them


'-----------------------------------
'designing and creating
'-----------------------------------

'Class Module ~ Film

option explicit

public title as string 'this populates the intelisense of the nromal module on the f. object instance
public releasedate as date 'this populates the intelisense of the nromal module on the f. object instance
public Lengthminutes as integer 'this populates the intelisense of the nromal module on the f. object instance

private sub class_initialize()
	worksheets("Sheet1").select
end sub




'Normal Module ~ TestFilmClass 
	Sub UsingFilmClass()

		Dim f as Film
	
		set f = New Film

		f.title = "Test"
		f.releasedate = "date"
		f.lengthminutes = 121

 		f.addtolist

		f.createwordreport

		set f = nothing ' destroys the instance of the class
		'variable is also destroyed when it goes out of scope
		'variables go out of scope when the subroutine they're run in ends

	End Sub



'-----------------------------------
'class events
'-----------------------------------
'activated much like worksheet events or workbook events 
'dropdown -> select Class
'other dropdown select - > Initialize/Terminate

'when the code in the normal module creates a new instance of the class if you 
'have the class initialize event going it will run all of the code in the initialize
'sub before returning back to the normal module 

'-----------------------------------
'class fields
'-----------------------------------
'public and private variable with declared data types able to be referenced upon new instance 
'of the class but if private it will not be accessable, and references must be set 
'and if private, option will not appear in intellisense

'-----------------------------------
'defining properties
'-----------------------------------
'3 types of property statements ""Let, Get, & Set""



'Let: allows you to pass a value into the property
'Get: allows you to return values from the property
'Set: when you want to pass in an object for your properties value

public property Let Title(Value as String) 'what we want the property to be called
	pTitle = Value 'read
end property
public property get title() as string
	title = pTitle 'write ability
end property

'properties are useful because you can pass the variable into the property like the film
'name and then a property can act as an independant subroutine and do things like
'validate that a copy doesnt exist already in the your data set, change the case, font,
'concatenate, etc.




'-----------------------------------
'creating methods 
'-----------------------------------
'basically a normal sub routine but when using the class instance f.AddToList
'the method will appear valid and the instance of the class will then perform the action of 
'the method subroutine


public sub AddToList()

	sheet1.select
	range("A1").end(xldown).offset(1,0).select 'select last cell in a then get next blank
	activecell.value=active.cell.offset(-1,0).value+1
	activecell.offset(0,1).value = me.Title 'Me is the instance of the class
	activecell.offset(0,2).value = me.ReleaseDate
	activecell.offset(0,3).value = me.LengthMinutes
	activecell.offset(0,4).value = me.GenreText
End Sub"
