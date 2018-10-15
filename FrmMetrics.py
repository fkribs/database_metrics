import clr
import re
from datetime import datetime, date

clr.AddReference('DB_Visualization_GUI.dll')
from DB_Visualization_GUI import FrmMetrics

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.IO')
from System.Windows.Forms import ListBox, Label, TableLayoutColumnStyleCollection
from System.Drawing import Size, Font, FontStyle

from System import DBNull, String
import math


class FrmMetrics(FrmMetrics):

    def __init__(self, rows, columns, metrics, selectedAttributes):
        self.minimums = []
        self.maximums = []

        self.rows = rows
        self.columns = columns
        self.columnNames = ["Attribute"] + metrics
        self.attributes = selectedAttributes

        self._initializeUserInterface()
        self._initializeBusinessLogic()

    def _initializeBusinessLogic(self):
        self._assignEventMethods()

    def _initializeUserInterface(self):
        self.Size = Size(1600, 500)
        self.tlpMain.ColumnCount = len(self.columnNames)
        self.tlpMain.RowCount = 2
        self.tlpMain.Size = self.ClientRectangle.Size
        #create list titles
        for name in self.columnNames:
            lbl = Label()
            lbl.Text = name
            lbl.Font = Font("Microsoft Sans Serif", 12, FontStyle.Bold)
            lbl.Height = 50
            self.tlpMain.Controls.Add(lbl)
        #create lists 
        for name in self.columnNames:
            lst = ListBox()
            lst.Name = name
            lst.Font = Font("Microsoft Sans Serif", 10)
            lst.Width = 135
            if name.ToLower() == 'attribute':
                lst.Width = 300
            if name.ToLower() == 'mode':
                lst.Width = 215
            lst.Height = self.ClientRectangle.Height - (lbl.Height * 2)
            self.tlpMain.Controls.Add(lst)
            self._populateLists(lst)

    def _assignEventMethods(self):
        #loop through table layout panel to modify each list
        for lst in self.tlpMain.Controls:
            #quick and dirty way to remove labels
            if lst.Name in self.columnNames:
                lst.SelectedIndexChanged += self._syncListsIndexes

    def _syncListsIndexes(self, sender, event):
         for lst in self.tlpMain.Controls:
            #quick and dirty way to remove labels
            if lst.Name in self.columnNames:
                try:
                    lst.SelectedIndex = sender.SelectedIndex
                except:
                    print("'{}' index is empty.".format(lst.Name))

    #identifies name of list and passes list to appropriate populate method
    def _populateLists(self, lst):
        name = lst.Name.ToLower()
        #populate attributes list
        if name == 'attribute':
            self._populateAttributes(lst)

        #populate minimums list
        if name == 'minimum':
            self.minimums = self._getMinimums() #minumums is in declared in module-level scope so that it can be passed to _getRanges as parameter
            for item in self.minimums:
                lst.Items.Add(item)

        #populate maximums list
        if name == 'maximum':
            self.maximums = self._getMaximums() #maximums is in declared in module-level scope so that it can be passed to _getRanges as parameter
            for item in self.maximums:
                lst.Items.Add(item)

        #populate ranges list
        if name == 'range':
            ranges = self._getRanges(self.minimums, self.maximums)
            for item in ranges:
                lst.Items.Add(item)

        #populate means list
        if name == 'mean':
            self.means = self._getMeans()
            for item in self.means:
                try:
                    lst.Items.Add('{0:.2f}'.format(item))
                except:
                    lst.Items.Add(item)

        #populate variance list
        if name == 'variance':
            self.variances = self._getVariances(self.means)
            for item in self.variances:
                try:
                    lst.Items.Add('{0:.2f}'.format(item))
                except:
                    lst.Items.Add(item)

        #populate standard deviation list
        if name == 'standard deviation':
            stdDevs = self._getStandardDeviations(self.variances)
            for item in stdDevs:
                try:
                    lst.Items.Add('{0:.2f}'.format(item))
                except:
                    lst.Items.Add(item)

        #populate mode list
        if name == 'mode':
            modes = self._getModesJSON()
            #print(modes)
            for attr in self.attributes:
                try:
                    modeVal = modes[attr][0]
                except:
                    modeVal = 'NA'
                if isinstance(modeVal, list):
                    try:
                        modeVal = modeVal[0]
                    except:
                        pass
                if isinstance(modeVal, DBNull):
                    modeVal = "DBNULL"
                    
                mode = modes[attr][1]
                lst.Items.Add("'{}': {}".format(modeVal, mode))


        #populate null count list
        if name == 'null count':
            nullCounts = self._getNullCountJSON()
            for attr in self.attributes:
                try:
                    nullCount = nullCounts[attr]
                except:
                    nullCount = 0
                total = nullCounts[attr + "total"] 
                percent = (float(nullCount) / float(total)) * 100
                
                formatted = "{0}/ {1} | {2:.2f}%".format(nullCount, total, percent)
                lst.Items.Add(formatted)

    def _populateAttributes(self,lst):
            for attribute in self.attributes:
                lst.Items.Add(attribute)

    def _getMinimums(self):
        #calculate minimum value for each selected column in DB, becomes row in metric form
        minimums = []
        for col in self.columns:
            val = float("inf")
            for row in self.rows:
                contents = row[col]
                if re.match("\d{4}-\d{2}-\d{2}", contents.ToString()):
                    if type(val) == float:
                        val = datetime.max.date()
                    contents = datetime.strptime(contents[:10], "%Y-%m-%d").date()
                    
                if isinstance(contents, DBNull):
                    continue
                try:
                    if int(contents.strip('"')) < val:
                        val = contents
                except:
                    if contents < val:
                        val = contents
            try:
                if (col.DataType.ToString() == 'System.String' and not re.match("\d{4}-\d{2}-\d{2}",row[col].ToString())) or val == float("inf"):
                    val = 'NA'
            except:
                pass
            minimums.Add(val)
        return minimums

    def _getMaximums(self):
        #calculate maximum value for each selected column in DB, becomes row in metric form
        maximums = []
        for col in self.columns:
            val = -1.0
            for row in self.rows:
                contents = row[col]
                if re.match("\d{4}-\d{2}-\d{2}", contents.ToString()):
                    if type(val) == float:
                        val = datetime.min.date()
                    contents = datetime.strptime(contents[:10], "%Y-%m-%d").date()

                if isinstance(contents, DBNull):
                    continue
                try:
                    if int(contents.strip('"')) > val:
                        val = contents
                except:
                    if contents > val:
                        val = contents
            if (col.DataType.ToString() == 'System.String' and not re.match("\d{4}-\d{2}-\d{2}",row[col].ToString())) or val == -1:
                val = 'NA'
            maximums.Add(val)
        return maximums

    def _getRanges(self, mins, maxs):
        #calculate range values between two lists for each index, return new list of ranges
        ranges = []
        for i in range(len(mins)):
            try:
                ranges.Add(maxs[i] - mins[i])
            except:
                ranges.Add('NA')
        return ranges

    def _getMeans(self):
        means = []
        for col in self.columns:
            size = 0.0
            sum = 0.0
            for row in self.rows:
                size += 1
                contents = row[col]
                if type(contents) != bool:
                    try:
                        sum += contents
                    except:
                        size -= 1
                else:
                    size -= 1
            try:
                means.Add(sum / size)
            except:
                means.Add('NA')
        return means


    def _getVariances(self,means):
        variances = []
        colIndex = 0
        currentVal = 0.0
        for col in self.columns:
            mean = self.means[colIndex]
            colIndex += 1
            rowCount = 0
            sumOfDifferenceSquared = 0.0
            for row in self.rows:
                currentVal = row[col]
                rowCount += 1
                try:
                    sumOfDifferenceSquared += (mean - currentVal)**2
                except:
                    rowCount -= 1
            try:
                variances.Add(sumOfDifferenceSquared / rowCount)
            except:
                variances.Add('NA')
        return variances


    def _getStandardDeviations(self,variances):
        sqrts = []
        for variance in variances:
            try:
                sqrts.Add(math.sqrt(variance))
            except:
                sqrts.Add('NA')
        return sqrts


    def _getModesJSON(self):
        valueCounts = {}
        modes = {}
        for col in self.columns:
            valueCounts[col.ColumnName] = {}
            rowNumber = 0
            for row in self.rows:
                value = row[col]
                try:
                    valueCounts[col.ColumnName][value] += 1
                except KeyError:
                    valueCounts[col.ColumnName][value] = 1
        #find modes
        for key in valueCounts:
            try:
                mode = max(valueCounts[key].values())
            except:
                mode = 'NA'
            modeAttr = [k for k, v in valueCounts[key].items() if v == mode]
            if len(modeAttr) >= 2:
                try:
                    modeAttr = modeAttr[0]
                except:
                    pass

            modes[key] = (modeAttr, mode)
        return modes
               


    def _getNullCountJSON(self):
            nulls = {}
            for col in self.columns:
                rowCount = 0
                for row in self.rows:
                    rowCount += 1
                    contents = row[col]
                    if isinstance(contents, DBNull):
                        try:
                            nulls[col.ColumnName] += 1
                        except:
                            nulls[col.ColumnName] = 1
                nulls[str(col.ColumnName) + "total"] = rowCount
            return nulls




