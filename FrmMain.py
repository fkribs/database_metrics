import clr
import System
import os
import re

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.IO')
from System.Windows.Forms import Application, RadioButton, TextBox, MessageBox, Label
from System.Drawing import Size, Font

clr.AddReference('DB_Visualization_GUI.dll')
from DB_Visualization_GUI import FrmMain 
from FrmMetrics import FrmMetrics

from ClsDB_Driver import *
from ClsUtilities import Util



class FrmMain(FrmMain):

    def __init__(self):
        self.productNames = ["TdpiBh3", "TdpiT1", "TdpiSc5", "TdpiS2", "TdpiQ", "WpiBh3", "WpiS2", "WpiSc5", "WpiT1", "TdpiD", "Ch"]
        self._initializeBusinessLogic()
        self._initializeUserInterface()
                
    def _initializeBusinessLogic(self):
        self.conn = Connector()
        self.util = Util()
        self._assignEventMethods()

    def _initializeUserInterface(self):
        self.Text = "DB Reporting"
        self._populateProductTypeGrpKey()
        self.grpTable.Controls[1].Checked = True
        self.grpProductFilterTabSelectTbcMain.Controls[0].Checked = True
        self.folderDialog.SelectedPath= os.getcwd()
        self.txtFilePath.Text = self.folderDialog.SelectedPath
        
    #populate listboxes with attributes of tables selected by radio button
    def _populateListBoxes(self, sender, event):
        tableName = ""
        if not sender.Checked:
            return
        #clear listboxes
        self.lstAvailableGrpAttributesTabSelectTbcMain.Items.Clear()
        self.lstSelectedGrpAttributesTabSelectTbcMain.Items.Clear()
        self.lstAvailableGrpAttributesTabWhereTbcMain.Items.Clear()
        self.lstSelectedGrpAttributesTabWhereTbcMain.Items.Clear()
        #get attributes from selected table
        tableName = sender.Text
        attributes = self.conn.getColumnNames(tableName)
        #filter attributes according to desired producttype
        attributes = self._removeAttributesByType(attributes)
        #populate listboxes with attributes
        [self.lstAvailableGrpAttributesTabSelectTbcMain.Items.Add(attr) for attr in attributes]
        [self.lstAvailableGrpAttributesTabWhereTbcMain.Items.Add(attr) for attr in attributes]

    #receives list of attributes and filters out and removes the
    def _removeAttributesByType(self,attributesList):
        numProducts = len(self.productNames)
        products = {i: {"name":"", "attributes":[]} for i in range(numProducts)}
        names = self.productNames
        
        #these are attributes that are consistent across all nodes
        genericAttrs = ["AvailableFwVersion", "BandgapInRange", "BandgapValue", "CurrentFwVersion", "Id", "DateTested", "ProductType", "IsPotted", "MoistureVInRange", "MoistureVValue", "PcbRevision", "SerialNumber", "PassedTest", "TemperatureVInRange", "TemperatureVValue", "Tester", "TesterMessage"] 
        for i in range(len(products)):
            attributes = []
            products[i]["name"] = names[i]

            #assign product specific attributes to products
            #"TdpiBh3", "TdpiT1", "TdpiSc5", "TdpiS2", "TdpiQ", "WpiBh3", "WpiS2", "WpiSc5", "WpiT1", "TdpiD", "Ch"
            if i == 0:
                attributes = genericAttrs + ["EnumAndAuthInRange","ExternalPowerVInRange", "ExternalPowerVValue", "Gen7BhSpeakerOk", "HeadNodeSpeakerOk", "SFlashIdOk", "SFlashReadingOk", "TwelveVPowerVInRange", "TwelveVPowerVValue", "VoltageAInRange", "VoltageAValue", "VoltageBInRange", "VoltageBValue"]
            if i == 1:
                attributes = genericAttrs + ["EnumAndAuthInRange","ExternalPowerVInRange", "ExternalPowerVValue", "Gen7BhSpeakerOk", "HeadNodeSpeakerOk", "SFlashIdOk", "SFlashReadingOk", "TwelveVPowerVInRange", "TwelveVPowerVValue", "VoltageAInRange", "VoltageAValue", "VoltageBInRange", "VoltageBValue"]
            products[i]["attributes"] = attributes
        
        product = ""
        for rad in self.grpProductFilterTabSelectTbcMain.Controls:
            if rad.Checked:
                product = rad.Text.lower()
        #this whole part might be unnecessary, could just return the attributes list created by the logic above
        #HOWEVER, the logic below guarantees that attributes returned were present in the parameter list.
        #The method is removeAttributesByType, not getAttributesByType. Perhaps this whole method is flawed by design?
        productNumber = -1
        newAttributes = []
        for i in products:
            if products[i]["name"].lower() == product:
                productNumber = i
        if productNumber == -1:
            return attributesList
        for attr in attributesList:
            if attr in products[productNumber]["attributes"]:
                newAttributes.append(attr)

        return newAttributes
    
    
    def _swapListBoxItems(self, sender, event):
        recipient = None
        #get text of tab that listbox is located in
        parentName = sender.Parent.Parent.Text
        #infer recipient
        if parentName == 'Select':
            if "available" in sender.Name.lower():
                recipient = self.lstSelectedGrpAttributesTabSelectTbcMain
            if "selected" in sender.Name.lower():
                recipient = self.lstAvailableGrpAttributesTabSelectTbcMain
                selectedList = True
        if parentName == 'Where':
            if "available" in sender.Name.lower():
                recipient = self.lstSelectedGrpAttributesTabWhereTbcMain
            if "selected" in sender.Name.lower():
                recipient = self.lstAvailableGrpAttributesTabWhereTbcMain
                selectedList = True
        #add sender SelectedItem to recipient Items
        if recipient:
            recipient.Items.Add(sender.SelectedItem)
        else:
            raise ValueError('Recipient listbox not assigned. It is likely that a tab or listbox was renamed.')

        #remove SelectedItem from sender list
        try: #wrapped in 'try' as quick fix for 'recipient.Items.Add()' triggering second SelectedIndexChanged event 
            sender.Items.Remove(sender.SelectedItem)
        except:
            pass
        self._alphabetizeListBox(recipient)
        #add where value textbox inputs
        self.flpValsTabWhereTbcMain.Controls.Clear()
        for item in self.lstSelectedGrpAttributesTabWhereTbcMain.Items:
            txtBox = TextBox()
            txtBox.MaximumSize = Size(170,19)
            txtBox.Width = 150
            txtBox.Font = Font("Microsoft Sans Serif", 7.25)
            txtBox.Text = item
            txtBox.Name = item
            txtBox.TextChanged += self._constructSQL
            self.flpValsTabWhereTbcMain.Controls.Add(txtBox)
    
    #set events to trigger methods
    def _assignEventMethods(self):
        #link radio buttons CheckChanged to _populateListBoxes method and _constructSQL method
        methods = [self._populateListBoxes, self._constructSQL]
        for method in methods:
            for radio in self.grpTable.Controls:
                radio.CheckedChanged += method
        #link list boxes SelectedIndexChanged to _swapListBoxItems method and _constructSQL method
        methods = [self._swapListBoxItems, self._constructSQL]
        listBoxes = [self.lstAvailableGrpAttributesTabSelectTbcMain,
                     self.lstSelectedGrpAttributesTabSelectTbcMain,
                     self.lstAvailableGrpAttributesTabWhereTbcMain,
                     self.lstSelectedGrpAttributesTabWhereTbcMain]
        for method in methods:
            for lb in listBoxes:
                lb.SelectedIndexChanged += method
        #link btnDisplayMetrics to executeQuery method
        self.btnDisplayMetrics.Click += self._generateMetrics
        self.btnSaveCSV.Click += self._saveCSV
        #link btnSet to openFileDialog
        self.btnSet.Click += self._openFolderDialog
        #link radio buttons CheckChanged to _populateListBoxes method
        for radio in self.grpProductFilterTabSelectTbcMain.Controls:
            radio.CheckedChanged += self._triggerGrpTableCheckedChanged

    #this is a workaround. Long story short, it's not possible to trigger _populateListBoxes directly from grpProductFilter
    #this is because _populateListBoxes performs logic with sender
    def _triggerGrpTableCheckedChanged(self, sender, event):
        for rad in self.grpTable.Controls:
            if rad.Checked:
                rad.Checked = False
                rad.Checked = True


    def _populateProductTypeGrpKey(self):
        self.tlpKeyGrpProductTypeKey.RowCount = len(self.productNames)
        i = -1
        for name in self.productNames:
            i+=1
            #must put str values into control because of TableLayoutPanel
            lblProduct = Label()
            lblProduct.Text = name
            lblNumber = Label()
            lblNumber.Text = str(i)
            self.tlpKeyGrpProductTypeKey.Controls.Add(lblProduct)
            self.tlpKeyGrpProductTypeKey.Controls.Add(lblNumber)


    #reorder listbox using Util class to handle logic (setting ID to top element)
    def _alphabetizeListBox(self, listBox):
        newList = self.util.sortObjectCollection(listBox.Items)
        listBox.Items.Clear()
        for item in newList:
            listBox.Items.Add(item)


    def _constructSQL(self, sender, event):
        #get table from selected radiobutton in Table groupbox
        table = "".join([radio.Text for radio in self.grpTable.Controls if radio.Checked])
        #get selected 'select' attributes
        selectAttributes = self.lstSelectedGrpAttributesTabSelectTbcMain.Items
        #construct 'select' string, reusing objectcollection variable
        if not selectAttributes:
            selectAttributes = '*'
        else:
            selectAttributes = "".join([item + ", " for item in selectAttributes]).Trim().Trim(",")
        #get selected 'where' attributes
        whereAttributes = self.lstSelectedGrpAttributesTabWhereTbcMain.Items
        #construct optional 'where' string if it exists, reusing objectcollection variable
        if not whereAttributes:
            whereAttributes = ";"
        else:
            #gets where values from generated text boxes.
            whereVals = {item:"= '" + self.flpValsTabWhereTbcMain.Controls.Find(item, False)[0].Text for item in whereAttributes}
            for key, val in whereVals.items():
                if "= '" + key == val:
                    #whereVals.pop(key,None)
                    #whereAttributes[whereAttributes.index(key)] = None
                    pass

            whereAttributes = "WHERE {}".format("".join([item + " {}'\r\n\tAND ".format(whereVals[item]) for item in whereAttributes if item != whereVals[item].Trim('=').Trim().Trim("'")]).Trim())
            if whereAttributes == 'WHERE ':
                whereAttributes = ""
            #remove last 'AND'
            if whereAttributes[-4:] == '\tAND':
                whereAttributes = "\r\n" + whereAttributes[:-4].Trim('\r').Trim('\n') + ';'
            else:
                whereAttributes = whereAttributes.Trim('\r').Trim('\n') + ';'
        #construct query
        self.query = "SELECT {} \r\nFROM {}{}".format(selectAttributes, table, whereAttributes)
        #display query in textbox
        self.txtTest.Text = self.query


    def _generateMetrics(self, sender, event):
        try:
            results = self.conn.executeGeneratedQuery(self.query)
        except:
            MessageBox.Show("Make sure your 'Where' datatypes are correct.", "Query Failed" )
            return
            
        try:
            rows = [row for row in results.Tables[0].Rows]
            columns = [col for col in results.Tables[0].Columns]
        except:
            rows = []
            columns = []
        metrics = ['Minimum', 'Maximum', 'Range', 'Mean', 'Variance','Standard Deviation', 'Mode', 'Null Count']

        #generalize dates so that they match more often
        for row in rows:
            try:
                if re.match("\d{4}-\d{2}-\d{2}",row[0].ToString()): 
                    for col in columns:
                        row[col] = row[col][:10]
            except:
                pass
        
        #create list of selected attributes
        selectedAttributes = [item for item in self.lstSelectedGrpAttributesTabSelectTbcMain.Items]
        if len(selectedAttributes) == 0:
            selectedAttributes = [item for item in self.lstAvailableGrpAttributesTabSelectTbcMain.Items]

        #show metrics form
        frmMetrics = FrmMetrics(rows, columns, metrics, selectedAttributes)
        frmMetrics.Show()


    def _saveCSV(self,sender,event):
        csv = ""
        results = self.conn.executeGeneratedQuery(self.query)
        rows = [row for row in results.Tables[0].Rows]
        columns = [col for col in results.Tables[0].Columns]

        csv = ", ".join([item for item in self.lstSelectedGrpAttributesTabSelectTbcMain.Items])
        if len(csv) == 0:
            csv = ", ".join([item for item in self.lstAvailableGrpAttributesTabSelectTbcMain.Items])

        for row in rows:
            csv += "\n"
            for col in columns:
                csv += "{}, ".format(row[col])

        
        fileName = "query_results"
        if self.txtFileName.Text != "":
            fileName = self.txtFileName.Text.replace(".csv",'')

        fileName = fileName + ".csv"
        filePath = os.path.join(self.txtFilePath.Text, fileName)
        try:
            with open(filePath,'w') as f:
                f.write(csv)
        except:
            status = "Save Failed."
            message = "Encountered an error while saving CSV. Ensure that file path is correct and not already open."
        status = "Save Complete."
        message = "CSV Saved Successfully to :\n\t{}".format(filePath)
        MessageBox.Show(message, status)


    def _openFolderDialog(self, sender, event):

        result = self.folderDialog.ShowDialog()
        path = ""
        if str(result) == "OK":
            path = self.folderDialog.SelectedPath
        else:
            path = os.getcwd()
        self.txtFilePath.Text = path


    
class FrmController():
    def start(self):
        frm = FrmMain()
        Application.Run(frm)
        return frm
