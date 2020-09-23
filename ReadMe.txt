MS Works 2000 Style Menus ReadMe
================================

Description: Replicates MS Works 2000 horizontal and verical menus


1.0.0 29/05/01
--------------
Initial Release

1.0.1 30/05/01
--------------
2 New Events: MouseEnter and MouseLeave
Added: Cursor support for both menus for seemless operation
Fixed: HoverItem not reset when mouse is moved too quickly
Fixed: HoverItem Event not generated for disabled MenuItems
Fixed: ShowHover Property did not Enable/Disable correctly (Thanks Thushan Fernando)
Fixed: MenuItem object only returned the SelectedItem instead of the requested Index
       Pointer from the collection.

1.0.2 7/06/01
-------------
Added ToolTip support

1.0.2a 8/06/01
--------------
Fixed: ucHMenu gets lost in an endless loop if no MenuItems are created. Problem
       does not exist for ucVMenu. (Thanks Dave Buckner)
Fixed: ucVMenu [HoverItem] encounters problems if no Menu Items exist (Thanks Dave 
       Buckner for finding and fixing)
Fixed: ucVMenu & ucHMenu Hover mode still worked even when disabled
Fixed: Tooltips constantly redrew which caused flickekering