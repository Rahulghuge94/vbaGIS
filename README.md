# vbaGIS
collection of class that provide GIS like functions in vba.

# usage
# point within polygon
# 'excel
select closed polyline in autocad then run macro from excel
Sub test3()
Dim acad As AcadApplication, doc As AcadDocument, pl As AcadLWPolyline, sel As AcadSelectionSet
Dim pt(0 To 1) As Double, tst As Variant
pt(0) = 463011.466: pt(1) = 2518727.857
Set acad = GetObject(, "AutoCAD.Application")
Set doc = acad.ActiveDocument
pt(0) = 463115
pt(1) = 2518668.6865
Set sel = doc.SelectionSets.Add("lrltf")
sel.Select acSelectionSetPrevious
For Each pl In sel
ReDim pol(0 To UBound(pl.Coordinates))
pol = varArray(pl.Coordinates)
Debug.Print pointInsidePolygon(pt, pol)
Next
sel.Delete
End Sub
# 'Autocad
select closed polyline in autocad then run macro from excel
Sub test3()
Dim pl As AcadLWPolyline, sel As AcadSelectionSet
Dim pt(0 To 1) As Double, tst As Variant
pt(0) = 463011.466: pt(1) = 2518727.857 'point to check if within polygon or not.
pt(0) = 463115
pt(1) = 2518668.6865
Set sel = ThisDrawing.SelectionSets.Add("lrltf")
sel.Select acSelectionSetPrevious
For Each pl In sel
ReDim pol(0 To UBound(pl.Coordinates))
pol = varArray(pl.Coordinates)
Debug.Print pointInsidePolygon(pt, pol)
Next
sel.Delete
End Sub
