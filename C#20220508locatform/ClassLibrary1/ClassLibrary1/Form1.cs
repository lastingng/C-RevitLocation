using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ClassLibrary1
{

    public partial class Form1 : System.Windows.Forms.Form
    {
        UIDocument uidoc;
        Document doc;
        Autodesk.Revit.UI.Selection.Selection sel;



        public Form1(UIDocument _uidoc, Document _doc, Autodesk.Revit.UI.Selection.Selection _sel)
        {
            InitializeComponent();
            uidoc = _uidoc;
            doc = _doc;
            sel = _sel;

            this.Load += Form1_Load;

        }

        private void Form1_Load(object sender, EventArgs E)
        {
            label1.Text = "S1:";
            label2.Text = "S2:";
            label3.Text = "Dim";
        }

        private void test_Click(object sender, EventArgs e)
        {
            foreach (ElementId eleid in sel.GetElementIds())
            {
                Element ele = doc.GetElement(eleid);
                FamilyInstance familyInstance = ele as FamilyInstance;

                if (familyInstance != null)
                {
                    FamilySymbol ftype = familyInstance.Symbol;
                    string famname = ftype.FamilyName;
                    string ang = "";
                    foreach (Parameter para in ele.Parameters)
                    {
                        if (para.Definition.Name == "Angle")
                        {
                            ang = para.AsValueString();
                        }
                    }

                    string convertAccName = this.converAccName(famname, ang);
                    label1.Text += convertAccName.ToString() + '\n';

                    foreach (Parameter para in ele.Parameters)
                    {
                        if (para.Definition.Name == "P-Section")
                        {
                            this.setParameter(para, doc, convertAccName);
                        }
                    }
                }
            }
        }

        private string converAccName(string famname, string angle)
        {
            if (famname.Contains("NRD"))
            {
                return "NRD";
            }
            else if (famname.Contains("Fire Damper"))
            {
                return "FD";
            }
            else if (famname.Contains("MFSD"))
            {
                return "MFSD";
            }
            else if (famname.Contains("Silencer"))
            {
                return "SIL";
            }
            else if (famname.Contains("Bend") | famname.Contains("Elbow"))
            {
                return angle + " Degree Elbow";
            }
            else
            {
                return "";
            }
        }

        private void setParameter(Autodesk.Revit.DB.Parameter para, Document doc, string tag)
        {
            Transaction transaction = new Transaction(doc);
            transaction.Start("para");
            para.Set(tag.ToString());
            transaction.Commit();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            IList<double> eleMinDimResult = new List<double>();
            double x = 0;
            double y = 0;
            double z = 0;

            foreach (ElementId eleid in sel.GetElementIds())
            {
                string i = textBox1.Text;
                Element ele = doc.GetElement(eleid);


                foreach (Parameter para1 in ele.Parameters)
                {
                    if (para1.Definition.Name == "P-Section")
                    {
                        if (para1.AsString() == i.ToString())
                        {

                            LocationPoint lp = ele.Location as LocationPoint;
                            LocationCurve lc = ele.Location as LocationCurve;

                            IList<ElementId> eleidDim = new List<ElementId>();
                            IList<double> ele1MinDim = new List<double>();
                            foreach (ElementId eleid1 in sel.GetElementIds())
                            {
                                Element ele1 = doc.GetElement(eleid1);


                                foreach (Parameter para in ele1.Parameters)
                                {

                                    if (para.Definition.Name == "P-Section")
                                    {

                                        if (eleid != eleid1 && para.AsString() == "")
                                        {
                                            LocationPoint lp1 = ele1.Location as LocationPoint;
                                            LocationCurve lc1 = ele1.Location as LocationCurve;


                                            if (lp != null)
                                            {
                                                x = lp.Point.X;
                                                y = lp.Point.Y;
                                                z = lp.Point.Z;

                                                double min = this.calEachPosition(lp1, lc1, x, y, z);
                                                eleidDim.Add(eleid1);
                                                ele1MinDim.Add(min);


                                            }
                                            else if (lc != null)
                                            {
                                                lc.Curve.GetEndPoint(0).ToString();
                                                x = lc.Curve.GetEndPoint(0).X;
                                                y = lc.Curve.GetEndPoint(0).Y;
                                                z = lc.Curve.GetEndPoint(0).Z;
                                                double minS = this.calEachPosition(lp1, lc1, x, y, z);
                                                eleidDim.Add(eleid1);
                                                ele1MinDim.Add(minS);

                                                x = lc.Curve.GetEndPoint(1).X;
                                                y = lc.Curve.GetEndPoint(1).Y;
                                                z = lc.Curve.GetEndPoint(1).Z;
                                                double minE = this.calEachPosition(lp1, lc1, x, y, z);
                                                eleidDim.Add(eleid1);
                                                ele1MinDim.Add(minE);

                                            }
                                        }
                                    }
                                }
                            }

                            if (eleid != null && eleidDim.Count != 0 && ele1MinDim.Count != 0 && ele1MinDim.Count != 0)
                            {
                                label1.Text += "\n" + eleid.ToString();
                                label2.Text += "\n" + eleidDim[Convert.ToInt32(ele1MinDim.IndexOf(ele1MinDim.Min()).ToString())].ToString();
                                ElementId elementidForSetPara = eleidDim[ele1MinDim.IndexOf(ele1MinDim.Min())];
                                Element elementForSetPara = doc.GetElement(elementidForSetPara);
                                foreach (Parameter para in elementForSetPara.Parameters)
                                {
                                    if (para.Definition.Name == "P-Section")
                                    {
                                        string paraSet = Convert.ToString(Convert.ToInt32(textBox1.Text) + 1);
                                        System.Threading.Thread.Sleep(500);
                                        this.setParameter(para, doc, paraSet);
                                    }

                                }
                                label3.Text += "\n" + ele1MinDim.Min().ToString();
                                textBox1.Text = Convert.ToString(Convert.ToInt32(textBox1.Text) + 1);
                                eleMinDimResult.Add(ele1MinDim.Min());
                            }
                        }
                    }
                }
            }
            label4.Text = "\n" + eleMinDimResult.Count.ToString();

        }
        private double calEachPosition(LocationPoint lp, LocationCurve lc, Double x, Double y, Double z)
        {
            IList<double> minResult = new List<double>();

            if (lp != null)
            {
                double x1 = lp.Point.X;
                double y1 = lp.Point.Y;
                double z1 = lp.Point.Z;

                double d = Math.Pow(Math.Pow((x - x1), 2) + Math.Pow((y - y1), 2) + Math.Pow((z - z1), 2), 0.5);

                minResult.Add(d);

            }
            else if (lc != null)
            {

                double x1 = lc.Curve.GetEndPoint(0).X;
                double y1 = lc.Curve.GetEndPoint(0).Y;
                double z1 = lc.Curve.GetEndPoint(0).Z;

                double d1 = Math.Pow(Math.Pow((x - x1), 2) + Math.Pow((y - y1), 2) + Math.Pow((z - z1), 2), 0.5);
                minResult.Add(d1);

                double x2 = lc.Curve.GetEndPoint(1).X;
                double y2 = lc.Curve.GetEndPoint(1).Y;
                double z2 = lc.Curve.GetEndPoint(1).Z;

                double d2 = Math.Pow(Math.Pow((x - x2), 2) + Math.Pow((y - y2), 2) + Math.Pow((z - z2), 2), 0.5);
                minResult.Add(d2);
            }

            return minResult.Min();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
        private string getParameterInformation(Autodesk.Revit.DB.Parameter para, Document doc)
        {
            switch (para.StorageType)
            {
                case StorageType.Double:
                    return para.AsValueString();

                case StorageType.Integer:
                    if (ParameterType.YesNo == para.Definition.ParameterType)
                    {
                        if (para.AsInteger() == 0)
                        {
                            return "false";
                        }
                        else
                        {
                            return "true";
                        }
                    }
                    else
                    {
                        return para.AsInteger().ToString();
                    }

                case StorageType.ElementId:
                    ElementId Id = para.AsElementId();
                    if (Id.IntegerValue >= 0)
                    {
                        return doc.GetElement(Id).Name;
                    }
                    else
                    {
                        return Id.IntegerValue.ToString();
                    }

                case StorageType.String:
                    return para.AsString();

                default:
                    return "Unknow Parameter";
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {



            ICollection<ElementId> eleColId = sel.GetElementIds();
            FilteredElementCollector FilterEleCol = new FilteredElementCollector(doc, eleColId);
            ElementCategoryFilter ElementFilter = new ElementCategoryFilter(BuiltInCategory.OST_DuctCurves);
            ElementCategoryFilter ElementFilter1 = new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory);
            LogicalOrFilter LogicFilter = new LogicalOrFilter(ElementFilter, ElementFilter1);


            ICollection<Element> ElementList = FilterEleCol.WherePasses(LogicFilter).WhereElementIsNotElementType().ToElements();

            foreach (Element element in ElementList)
            {

                label1.Text += "\n" + element.Category.Id.ToString();
            }



        }

        private void button2_Click(object sender, EventArgs e)
        {
            ICollection<ElementId> ElementId = sel.GetElementIds();
            foreach (ElementId eleid in ElementId)
            {
                Element element = doc.GetElement(eleid);
                BoundingBoxXYZ bbxyz = element.get_BoundingBox(null);

                Outline outline = new Outline(bbxyz.Min, bbxyz.Max);

                BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);

                IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();
                StringBuilder stringBuilder = new StringBuilder("this elment is " + element.Category.Name + "\n");
                stringBuilder.Append("-------------------------" + "\n");



                foreach (Element elementfilter in inter)
                {
                    LocationPoint Lp = elementfilter.Location as LocationPoint;
                    LocationCurve Lc = elementfilter.Location as LocationCurve;

                    stringBuilder.Append("Category : " + elementfilter.Category.Name + ", ElementID : " + elementfilter.Id.ToString() + "\n");
                    if (Lp != null)
                    {
                        stringBuilder.Append(Lp.Point.ToString() + "\n");
                    }
                    else if (Lc != null)
                    {
                        stringBuilder.Append(Lc.Curve.GetEndPoint(1).ToString() + " " + Lc.Curve.GetEndPoint(0).ToString() + "\n");
                    }
                }

                stringBuilder.Append("-------------------------" + "\n");

                label1.Text += "\n" + stringBuilder.ToString();

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            IList<Element> revitLink = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements();

            foreach (RevitLinkInstance link in revitLink)
            {
                Document LinkDoc = link.GetLinkDocument();

                using (FilteredElementCollector linkFoundations = new FilteredElementCollector(LinkDoc)
                    .OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType())
                {
                    foreach (Element element in linkFoundations.ToElements())
                    {
                        OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                        ogs = ogs.SetSurfaceTransparency(100);


                        using (Transaction trans = new Transaction(doc, "Foundation overrides"))
                        {
                            trans.Start();
                            try
                            {
                                doc.ActiveView.SetCategoryOverrides(element.Category.Id, ogs);
                                label1.Text += "\n" + element.Name.ToString();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("override error", ex.ToString());
                            }
                            trans.Commit();
                        }
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            Reference refElemLinked = uidoc.Selection.PickObject(ObjectType.LinkedElement, "Select edge");
            RevitLinkInstance elem = doc.GetElement(refElemLinked.ElementId) as RevitLinkInstance;
            Document docLinked = elem.GetLinkDocument();

            Element linkedelement = docLinked.GetElement(refElemLinked.LinkedElementId);

            LocationPoint Lp = linkedelement.Location as LocationPoint;
            LocationCurve Lc = linkedelement.Location as LocationCurve;
            if (Lp != null)
            {
                label1.Text += "\n" + linkedelement.Name;
                label2.Text += "\n" + Lp.Point.ToString();
            }
            else if (Lc != null)
            {
                label1.Text += "\n" + linkedelement.Name;
                label2.Text += "\n" + Lc.Curve.GetEndPoint(1).ToString() + " " + Lc.Curve.GetEndPoint(0).ToString();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {


            ICollection<ElementId> ElementId = sel.GetElementIds();

            foreach (ElementId eleid in ElementId)
            {
                Element ele = doc.GetElement(eleid);
                LocationPoint Lp = ele.Location as LocationPoint;
                LocationCurve Lc = ele.Location as LocationCurve;
                if (Lp != null)
                {
                    label1.Text += "\n" + ele.Name;
                    label2.Text += "\n" + Lp.Point.ToString();
                }
                else if (Lc != null)
                {
                    label1.Text += "\n" + ele.Name;
                    label2.Text += "\n" + Lc.Curve.GetEndPoint(1).ToString() + " " + Lc.Curve.GetEndPoint(0).ToString();
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));

            foreach (Element ele in collector)
            {
                label1.Text += "\n" + ele.Name;
            }
        }
    }
}
