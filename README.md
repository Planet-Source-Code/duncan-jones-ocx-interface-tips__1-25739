<div align="center">

## OCX interface tips


</div>

### Description

3 tips that you should definitely use to make your own developed ActiveX controls more developer friendly:
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-jones.md)
**Level**          |Intermediate
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-jones-ocx-interface-tips__1-25739/archive/master.zip)





### Source Code

<P><B>(1) Use enumerated types wherever you can in your control's interface</P>
- When you create a property for your control, try to define it as an enumerated type wherever possible. Doing so allows the Visual Basic IDE to populate a listbox with the values and their description in the "properties" tab. Thus where the ImageMap control has a ControlMode property which says how the control is behaving, the property is implemented as an enumerated type:
</P>
<P>
<PRE>
Private Enum enControlMode
  cm_Runtime = 1     '\\ The control is in runtime mode
  cm_DesignTime_Selecting '\\ Control is design time but no ImageMapItem is selected
  cm_DesignTime_Selected '\\ Control is in design time and an ImageMapItem is selected
  cm_DesignTime_Drawing  '\\ Control is drawing a new ImageMapItem
End Enum
</PRE>
</P>
<P>
This gives the developer more information about what the property is all about.
</P>
<P><B>(2) Use the "Procedure Attributes" screen to give extra information about your properties.</B> Bringing up the procedure attributes screen to set extra information about your properties. In particular, pressing the "Advanced >>>" button brings up three drop down boxes - Procedure ID:, Use this page in property browser:, and Property Category:.</P>
<P><B>(3) Fill the "Description" box of the "Procedure Attributes" box with a short but informative description.</B> This information is displayed when that property is the currently selected one in the property browser widnow.
</P>

