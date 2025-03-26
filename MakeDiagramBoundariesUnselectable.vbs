option explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: MakeDiagramBoundariesUnselectable
' Purpose: Makes all currently selectable Boundary elements in the diagram unselectable
' Author: JCH with Grok AI assist
' Date: March 25, 2025

sub main
    dim diagram
    dim diagramObject
    dim element
    dim updated
    dim boundaryCount
    dim changedCount
    updated = false
    boundaryCount = 0
    changedCount = 0
    
    ' Try to get the current diagram
    set diagram = Repository.GetCurrentDiagram
    
    if diagram is nothing then
        Session.Output "No diagram is currently open. Attempting to select one..."
        set diagram = Repository.GetTreeSelectedObject
        if not diagram is nothing and diagram.ObjectType = otDiagram then
            Session.Output "Selected diagram: " & diagram.Name & " (ID: " & diagram.DiagramID & ")"
        else
            Session.Output "Please open or select a diagram and try again."
            exit sub
        end if
    else
        Session.Output "Diagram found: " & diagram.Name & " (ID: " & diagram.DiagramID & ")"
    end if
    
    ' Loop through all objects in the diagram
    for each diagramObject in diagram.DiagramObjects
        set element = Repository.GetElementByID(diagramObject.ElementID)
        if not element is nothing then
            Session.Output "Element found - Type: " & element.Type & ", Name: " & element.Name & ", ID: " & element.ElementID
            if UCase(element.Type) = "BOUNDARY" then
                boundaryCount = boundaryCount + 1
                Session.Output "Boundary detected - Current IsSelectable: " & diagramObject.IsSelectable
                if diagramObject.IsSelectable then  ' Only update if currently selectable
                    diagramObject.IsSelectable = False
                    if diagramObject.Update then
                        Session.Output "Boundary updated successfully - New IsSelectable: " & diagramObject.IsSelectable
                        updated = true
                        changedCount = changedCount + 1
                    else
                        Session.Output "Failed to update boundary (ID: " & element.ElementID & ")"
                    end if
                else
                    Session.Output "Boundary already unselectable - No change needed"
                end if
            end if
        end if
    next
    
    ' Refresh the diagram if changes were made
    if updated then
        Repository.SaveDiagram diagram.DiagramID
        Repository.CloseDiagram diagram.DiagramID  ' Force close
        Repository.OpenDiagram diagram.DiagramID   ' Reopen to refresh
        Session.Output "Diagram fully refreshed. Total boundaries made unselectable: " & changedCount & " (out of " & boundaryCount & " boundaries found)"
    else
        Session.Output "No selectable boundaries found to update. Total boundaries: " & boundaryCount
    end if
end sub

main
