/*
Create a dimension line between all chosen grid. The gird lines need to be parallel to each other,
which is the case most of the time but not always.
TESTED REVIT API: 2018
The snippet can be used as is in a Revit Application Macro for test purposes
Author: Deyan Nenov | github.com/ArchilizerLtd | www.archilizer.com 
This file is shared on www.revitapidocs.com
For more information visit http://github.com/gtalarico/revitapidocs
License: http://github.com/gtalarico/revitapidocs/blob/master/LICENSE.md
*/
/// 

/// Creates a single dimension string between the chosen Grid lines
/// 

public void DimGrids()
{			
  UIDocument uidoc = this.ActiveUIDocument;
  Document doc = uidoc.Document;

  // Pick all the grid lines you want to dimension to
  GridSelectionFilter filter = new ThisApplication.GridSelectionFilter(doc);			
  var grids = uidoc.Selection.PickElementsByRectangle(filter, "Pick Grid Lines");

  ReferenceArray refArray = new ReferenceArray();
  XYZ dir = null;

  foreach(Element el in grids)
  {
    Grid gr = el as Grid;

    if(gr == null) continue;	
    if(dir == null)
    {
      Curve crv = gr.Curve;
      dir = new XYZ(0,0,1).CrossProduct((crv.GetEndPoint(0) - crv.GetEndPoint(1)));	// Get the direction of the gridline
    }

    Reference gridRef = null;

    // Options to extract the reference geometry needed for the NewDimension method
    Options opt = new Options();
    opt.ComputeReferences = true;
    opt.IncludeNonVisibleObjects = true;
    opt.View = doc.ActiveView;
    foreach (GeometryObject obj in gr.get_Geometry(opt))
    {
      if (obj is Line)
      {
        Line l = obj as Line;
        gridRef = l.Reference;
        refArray.Append(gridRef);	// Append to the list of all reference lines 
      }
    }
  }

  XYZ pickPoint = uidoc.Selection.PickPoint();	// Pick a placement point for the dimension line
  Line line = Line.CreateBound(pickPoint, pickPoint + dir * 100);		// Creates the line to be used for the dimension line

  using(Transaction t = new Transaction(doc, "Make Dim"))
  {
    t.Start();
    if( !doc.IsFamilyDocument )
    {
      doc.Create.NewDimension( 
        doc.ActiveView, line, refArray);
    }				
    t.Commit();
  }					
}
/// 

/// Grid Selection Filter (example for selection filters)
/// 

public class GridSelectionFilter : ISelectionFilter
{
  Document doc = null;
  public GridSelectionFilter(Document document)
  {
    doc = document;
  }

  public bool AllowElement(Element element)
  {
    if(element.Category.Name == "Grids")
    {
      return true;
    }
    return false;
  }

  public bool AllowReference(Reference refer, XYZ point)
  {
    return true;
  }
}