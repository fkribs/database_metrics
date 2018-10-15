import clr
clr.AddReference('System.Data')
from System.Data import DataSet
from System.Data.Odbc import OdbcConnection, OdbcDataAdapter
from ClsCredentials import ClsCredentials


class Connector():

    def __init__(self, connectString=ClsCredentials.CONNECTSTRING):
        self.connectString = connectString

    def _executeQuery(self, query):
        connection = OdbcConnection(self.connectString)
        adaptor = OdbcDataAdapter(query, connection)
        dataSet = DataSet()
        connection.Open()
        adaptor.Fill(dataSet)
        connection.Close()
        return dataSet

    def executeGeneratedQuery(self, query):
        return self._executeQuery(query)

    def getColumnNames(self, tableName):

        query = "SELECT * FROM {}".format(tableName)
        dataSet = self._executeQuery(query)
        columnNames = [column.ColumnName for column in dataSet.Tables[0].Columns]

        return columnNames


