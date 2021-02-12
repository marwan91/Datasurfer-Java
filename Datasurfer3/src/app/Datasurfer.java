package app;
import java.io.File;  
import java.io.FileInputStream;  
import java.util.Iterator;  
import java.util.ArrayList;
import java.awt.Color;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import javax.swing.JComboBox;


import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  


public class Datasurfer  
{  
	
static ArrayList<String> paths = new ArrayList<String>();
	
static ArrayList<String> description = new ArrayList<String>();
static ArrayList<String> pn = new ArrayList<String>();
static ArrayList<String> sn = new ArrayList<String>();
static ArrayList<String> loc = new ArrayList<String>();
static ArrayList<String> qty = new ArrayList<String>();
static ArrayList<String> method = new ArrayList<String>();
static ArrayList<String> sheet_names = new ArrayList<String>();
static ArrayList<String> sheet_method = new ArrayList<String>();
static ArrayList<String> sheet_key = new ArrayList<String>();
static ArrayList<String> local_sheet_key = new ArrayList<String>();

static int key_counter = 0;
static boolean ready = false;

static ArrayList<String> main_method_list = new ArrayList<String>();

private String[] methods = { "---", "Eddy Current", "Ultrasonic","Thermography", "X-Ray" };




public static void main(String[] args)   
{ 
	
	
	String[][] mydata = new String[5][2];
	
	
try  
{ 
	

///////////////  listing files //////////////
	
	File folder = new File("Capability\\eddycurrent");
	
	for (final File fileEntry : folder.listFiles()) {
        if (fileEntry.isDirectory()) {
            continue;
        } else {
            paths.add("Capability\\eddycurrent\\"+  fileEntry.getName());
            main_method_list.add("Eddy Current");
        }
    }
	
folder = new File("Capability\\thermography");
	
	for (final File fileEntry : folder.listFiles()) {
        if (fileEntry.isDirectory()) {
            continue;
        } else {
            paths.add("Capability\\thermography\\"+  fileEntry.getName());
            main_method_list.add("Thermography");
        }
    }
	
folder = new File("Capability\\xray");
	
	for (final File fileEntry : folder.listFiles()) {
        if (fileEntry.isDirectory()) {
            continue;
        } else {
            paths.add("Capability\\xray\\"+  fileEntry.getName());
            main_method_list.add("X-Ray");
        }
    }
	
folder = new File("Capability\\ultrasonic");
	
	for (final File fileEntry : folder.listFiles()) {
        if (fileEntry.isDirectory()) {
            continue;
        } else {
            paths.add("Capability\\ultrasonic\\"+  fileEntry.getName());
            main_method_list.add("Ultrasonic");
        }
    }
	
	for(int i=0;i<paths.size();i++ )
	{
		System.out.println(paths.get(i));
	}
	
/////////////////////////////////////////////
	
	
	
for(int j=0;j<paths.size();j++)
{
File file = new File(paths.get(j));   //creating a new file instance  
FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file    
XSSFWorkbook wb = new XSSFWorkbook(fis);  

for(int i=0;i<wb.getNumberOfSheets();i++ )
{
XSSFSheet sheet = wb.getSheetAt(i);     //creating a Sheet object to retrieve object  
Iterator<Row> itr = sheet.iterator();    //iterating over excel file 

sheet_names.add( wb.getSheetAt(i).getSheetName() );
sheet_method.add(main_method_list.get(j));

while (itr.hasNext())                 
{  
Row row = itr.next();

try {
	if(row.getCell(1).getCellType()== Cell.CELL_TYPE_STRING && row.getCell(3).getCellType()== Cell.CELL_TYPE_STRING && row.getCell(4).getCellType()== Cell.CELL_TYPE_STRING)
	{
	String cell1 = row.getCell(1).toString();
	String cell2 = row.getCell(3).toString();
	String cell3 = row.getCell(4).toString();
	
	if (cell1.contains("DESCRIPITON") && cell2.contains("P/N") && cell3.contains("S/N")) {;continue;}
	}
	else 
	{	
	if(row.getCell(1).getCellType()!= Cell.CELL_TYPE_STRING && row.getCell(3).getCellType()!= Cell.CELL_TYPE_STRING && row.getCell(4).getCellType()!= Cell.CELL_TYPE_STRING)
	{
		
		continue;}
	}
} catch (Exception e1) {
	continue;
}

if(row.getCell(1).toString()=="" && row.getCell(0).toString()!="")
{
description.add(row.getCell(0).toString());	
}
else 
{
description.add(row.getCell(1).toString());	
}

pn.add(row.getCell(3).toString());
sn.add(row.getCell(4).toString());
loc.add(row.getCell(6).toString());
qty.add(row.getCell(7).toString()); 
method.add(main_method_list.get(j));
sheet_key.add(  String.valueOf(key_counter));




Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
while (cellIterator.hasNext())   
{  
Cell cell = cellIterator.next();  
  
}  
  
} 
key_counter++;
}
}


mydata = new String[description.size()][6];
for(int i=0;i<description.size();i++)
{
	mydata[i][0]=description.get(i);
	mydata[i][1]=pn.get(i);
	mydata[i][2]=sn.get(i);
	mydata[i][3]=loc.get(i);
	mydata[i][4]=qty.get(i);
	mydata[i][5]=method.get(i);
}



}  
catch(Exception e)  
{  
e.printStackTrace();  
}  
////////////////////////////////   GUI  ///////////////////////////////////////

JFrame f=new JFrame("NDT Department Capability");  
final JTextField tf=new JTextField();  
tf.setBounds(20,30, 150,20);  
JButton b=new JButton("Search");  
b.setBounds(20,50,95,30);  
JComboBox c = new JComboBox();
c.setBounds(150,150,120,30);
c.setBackground(Color.white);
c.addItem("---");
c.addItem("Eddy Current");
c.addItem("Ultrasonic");
c.addItem("X-Ray");
c.addItem("Thermography");
c.setBackground(Color.white);
JComboBox c2 = new JComboBox();
c2.setBounds(450,150,120,30);
c2.addItem("---");
c2.setBackground(Color.white);


JLabel l1=new JLabel("Search By Part Number or Serial Number");  
l1.setBounds(10,10, 250,20); 
JLabel l2=new JLabel("Or Choose Method");
l2.setBounds(30,150, 250,20);
JLabel l3=new JLabel("Then Choose Location");
l3.setBounds(310,150, 250,20);

String data[][]= new String[1][6];  
final String column[]={"Description","P/N","S/N","Location","Quantity","Method"};         
JTable jt=new JTable(data,column);    
jt.setBounds(20,210,800,400);    
JScrollPane sp=new JScrollPane(jt); 
sp.setBounds(20, 220, 800, 400);



c.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent e) {
    	
    	ready = false;
        String selected_method =  ((JComboBox) e.getSource()).getSelectedItem().toString();
        c2.removeAllItems();
        c2.addItem("---");
        local_sheet_key.clear();
        c2.setSelectedIndex(0);
        
        for(int k=0;k<sheet_method.size();k++)
        {
        	if(selected_method==sheet_method.get(k)) {c2.addItem(sheet_names.get(k));local_sheet_key.add(String.valueOf(k));}
        }
       ready = true; 
    }
  });

c2.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent e) {
    	
    	if(!ready) {return;}
    	ArrayList<String> desc_result = new ArrayList<String>();
    	ArrayList<String> pn_result = new ArrayList<String>();
    	ArrayList<String> sn_result = new ArrayList<String>();
    	ArrayList<String> loc_result = new ArrayList<String>();
    	ArrayList<String> qty_result = new ArrayList<String>();
    	ArrayList<String> method_result = new ArrayList<String>();
    	
       int index = c2.getSelectedIndex()-1;
       if(c2.getSelectedItem()=="---") {return;}
       
   
       
       
       for( int j=0;j<description.size();j++)
       {
    	   
    	   if(  (local_sheet_key.get(index).equals(sheet_key.get(j) ))) 
    	   {
    		desc_result.add(description.get(j));
   			pn_result.add(pn.get(j));
   			sn_result.add(sn.get(j));
   			loc_result.add(loc.get(j));
   			qty_result.add(qty.get(j));
   			method_result.add(method.get(j));
    		   
    		   
    	   }
    	   	   
       }
       
    desc_result.add("");
   	pn_result.add("");
   	sn_result.add("");
   	loc_result.add("");
   	qty_result.add("");
   	method_result.add("");
   	
   	
String[][] result_array = new String[desc_result.size()][6];
	
	DefaultTableModel result_model = new DefaultTableModel();
	result_model.setColumnCount(6);
	result_model.setRowCount(desc_result.size());
	for(int i=0;i<desc_result.size();i++)
	{
		result_model.setValueAt(desc_result.get(i),i, 0);
		result_model.setValueAt(pn_result.get(i),i, 1);
		result_model.setValueAt(sn_result.get(i),i, 2);
		result_model.setValueAt(loc_result.get(i),i, 3);
		result_model.setValueAt(qty_result.get(i),i, 4);
		result_model.setValueAt(method_result.get(i),i, 5);
	}
	
	result_model.setColumnIdentifiers(column);
	
	
	jt.setModel(result_model); 
      
    }
  });






b.addActionListener(new ActionListener(){  
public void actionPerformed(ActionEvent e){ 
	///////////// Search Function//////
	
	ArrayList<String> desc_result = new ArrayList<String>();
	ArrayList<String> pn_result = new ArrayList<String>();
	ArrayList<String> sn_result = new ArrayList<String>();
	ArrayList<String> loc_result = new ArrayList<String>();
	ArrayList<String> qty_result = new ArrayList<String>();
	ArrayList<String> method_result = new ArrayList<String>();
	
	c.setSelectedIndex(0);
	c2.setSelectedIndex(0);
	if(tf.getText().isBlank()) {return;}
	
	
	for(int k=0;k<description.size();k++)
	{
		if(  pn.get(k).contains(tf.getText()) || sn.get(k).contains(tf.getText()) )
		{
			desc_result.add(description.get(k));
			pn_result.add(pn.get(k));
			sn_result.add(sn.get(k));
			loc_result.add(loc.get(k));
			qty_result.add(qty.get(k));
			method_result.add(method.get(k));
		}
		else
		{
		continue;
		}
			
	}
	
	desc_result.add("");
	pn_result.add("");
	sn_result.add("");
	loc_result.add("");
	qty_result.add("");
	method_result.add("");
	
	String[][] result_array = new String[desc_result.size()][6];
	
	DefaultTableModel result_model = new DefaultTableModel();
	result_model.setColumnCount(6);
	result_model.setRowCount(desc_result.size());
	for(int i=0;i<desc_result.size();i++)
	{
		result_model.setValueAt(desc_result.get(i),i, 0);
		result_model.setValueAt(pn_result.get(i),i, 1);
		result_model.setValueAt(sn_result.get(i),i, 2);
		result_model.setValueAt(loc_result.get(i),i, 3);
		result_model.setValueAt(qty_result.get(i),i, 4);
		result_model.setValueAt(method_result.get(i),i, 5);
	}
	
	result_model.setColumnIdentifiers(column);
	
	
	
	jt.setModel(result_model);      
	
    ///////////////////////////////////
    }  
});  


f.add(sp);
f.add(b);f.add(tf);f.add(l1);f.add(l2);f.add(l3);f.add(c);f.add(c2);
f.setSize(600,600);  
f.setLayout(null);  
f.setVisible(true); 

///////////////////////////////////////////////////////////////////////////////
}  
}  