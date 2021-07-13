@Grab('net.codebuilders:bootstrap-connector:4.1.6.0')
@Grab("net.codebuilders:juh:4.1.6")
@Grab("net.codebuilders:ridl:4.1.6")
@Grab("net.codebuilders:unoil:4.1.6")
@Grab("net.codebuilders:jurt:4.1.6")
@Grab('net.codebuilders:guno-extension:4.1.6.16')
@Grab('com.opencsv:opencsv:5.2')

import com.sun.star.beans.XPropertySet
import com.sun.star.comp.helper.BootstrapException
import com.sun.star.connection.NoConnectException
import com.sun.star.frame.XComponentLoader
import com.sun.star.frame.XController
import com.sun.star.frame.XModel
import com.sun.star.lang.XComponent
import com.sun.star.lang.XMultiComponentFactory
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.sheet.XSpreadsheets
import com.sun.star.sheet.XSpreadsheetView
import com.sun.star.sheet.XViewFreezable
import com.sun.star.table.*
import com.sun.star.uno.UnoRuntime
import com.sun.star.uno.XComponentContext

import ooo.connector.BootstrapSocketConnector
import groovy.time.TimeCategory
import groovy.time.TimeDuration
import groovy.transform.CompileStatic

import com.opencsv.bean.CsvToBeanBuilder


// @CompileStatic
class GroovyClient {

    // location of openoffice executable soffice
    static String oooExeFolder = "/opt/openoffice4/program"
    // static String oooExeFolder = "C:/Program Files (x86)/OpenOffice 4/program"

    static XComponentContext mxRemoteContext
    static XMultiComponentFactory mxRemoteServiceManager
    static XComponent xComponent
    static XSpreadsheetView xSpreadsheetView
    static XViewFreezable xFreeze

    // Colors are in Hex format RRGGBB from 0(00) to 255(ff)
    // Standard Colors
    // static final Integer RED = 0xff0000
    // static final Integer GREEN = 0x00ff00
    // static final Integer BLUE = 0x0000ff
    static final Integer BLACK = 0x000000
    static final Integer WHITE = 0xffffff

    // Custom Colors
    static final Integer BLUE = 0x004b8d
    static final Integer GREEN = 0x62bb46
    static final Integer TURQUOISE = 0x00a4d2
    static final Integer GOLD = 0xffcf22
    static final Integer ORANGE = 0xf79428
    static final Integer PURPLE = 0x6e298d
    static final Integer SKY_BLUE = 0x87e5ff
    static final Integer BLUE_GREEN = 0x00aa7e
    static final Integer LIGHT_GRAY = 0x959797
    static final Integer CRIMSON = 0xd31245
    static final Integer BROWN = 0x8a4b05
    static final Integer DARK_GRAY = 0x3f4040
    static final Integer LIGHTER_GRAY = 0xf2f2f2
    static final Integer LIGHT_YELLOW = 0xffff99


    GroovyClient() {

    }

    static void setupDocument(XSpreadsheetDocument doc) {

        try {
            XPropertySet colPs = doc.getCellStylePropertySet("ColHeading")
            // set properties
            colPs.putAt("IsCellBackgroundTransparent", false)
            colPs.putAt("CellBackColor", BLUE)
            colPs.putAt("CharColor", WHITE)
            colPs.putAt("IsTextWrapped", true)
            // colPs.putAt("RotateAngle", 9000) // angle * 100
            colPs.putAt("VertJustify", CellVertJustify.BOTTOM)
            colPs.putAt("HoriJustify", CellHoriJustify.CENTER)

            XPropertySet rowPs = doc.getCellStylePropertySet("RowHeading")
            // set properties
            rowPs.putAt("HoriJustify", CellHoriJustify.LEFT)

            XPropertySet oddRowPs = doc.getCellStylePropertySet("OddRow")
            // set properties
            oddRowPs.putAt("HoriJustify", CellHoriJustify.CENTER)
            oddRowPs.putAt("IsCellBackgroundTransparent", false)
            oddRowPs.putAt("CellBackColor", LIGHT_YELLOW)
            oddRowPs.putAt("CharColor", DARK_GRAY)

            XPropertySet evenRowPs = doc.getCellStylePropertySet("EvenRow")
            // set properties
            evenRowPs.putAt("HoriJustify", CellHoriJustify.CENTER)
            evenRowPs.putAt("IsCellBackgroundTransparent", false)
            evenRowPs.putAt("CellBackColor", LIGHTER_GRAY)
            evenRowPs.putAt("CharColor", DARK_GRAY)

            XPropertySet redBgPs = doc.getCellStylePropertySet("RedBg")
            // set properties
            // redBgPs.setPropertyValue("HoriJustify", CellHoriJustify.CENTER)
            redBgPs.putAt("IsCellBackgroundTransparent", false)
            redBgPs.putAt("CellBackColor", CRIMSON)
            redBgPs.putAt("CharColor", WHITE)


        } catch (Exception e) {
            e.printStackTrace(System.err);
        }

        try {
            // get the spreadsheet view (used to set active sheet)
            XModel xModel = doc.guno(XModel.class)
            XController xController = xModel.currentController
            // com.sun.star.sheet.XViewFreezable xFreeze
            xSpreadsheetView = xController.guno(XSpreadsheetView.class)
            xFreeze = xController.guno(XViewFreezable.class)

        } catch (Exception e) {
            System.out.println("Couldn't get SpreadsheetView " + e)
            e.printStackTrace(System.err)
        }

        println("Office Initialized...")

    }

    // http://opencsv.sourceforge.net
    // http://opencsv.sourceforge.net/apidocs/index.html
    static List<Employee> getEmployees(String fp) {

        List<Employee> employees = new CsvToBeanBuilder(new FileReader(fp))
                .withType(Employee.class).build().parse()

        return employees
    }

    static void formatSheet(XSpreadsheet sht, Integer empCount) {

        println("Creating the Header")
        // column, row, string, sheet, flag
        // rows begin at 0, columns at 0
        sht.setFormulaOfCell(0, 0, "Last Name")
        sht.setFormulaOfCell(1, 0, "First Name")
        sht.setFormulaOfCell(2, 0, "Start Date")
        sht.setFormulaOfCell(3, 0, "Email")


        int colCount = 4

        // set the cell style of the header
        XCellRange xCR = null
        xCR = sht.getCellRangeByPosition(0, 0, (colCount - 1), 0)

        XPropertySet xCPS = xCR.guno(XPropertySet.class)
        xCPS.putAt("CellStyle", "ColHeading")

        // Stripe the content rows
        (1..empCount).each { r ->
            xCR = null
            xCPS = null
            xCR = sht.getCellRangeByPosition(0, r, (colCount - 1), r)

            xCPS = xCR.guno(XPropertySet.class)

            if (r % 2) {
                // evenRowPs
                xCPS.putAt("CellStyle", "EvenRow")
            } else {
                // oddRowPs
                xCPS.putAt("CellStyle", "OddRow")
            }
        }

        // set 1st column width
        XCellRange xCellRange = null
        xCellRange = sht.getCellRangeByName("A1:D1")
        XColumnRowRange xColRowRange = xCellRange.guno(XColumnRowRange.class)
        XTableColumns xColumns = xColRowRange.columns

        Object aColumnObj = xColumns.getByIndex(2)
        XPropertySet aColPS = aColumnObj.guno(XPropertySet.class)
        aColPS.putAt("Width", 6000)

        aColumnObj = xColumns.getByIndex(3)
        aColPS = aColumnObj.guno(XPropertySet.class)
        aColPS.putAt("Width", 6000)

        XTableRows xRows = xColRowRange.rows
        Object aRowObj = xRows.getByIndex(0)
        XPropertySet firstRowPS = aRowObj.guno(XPropertySet.class)
        firstRowPS.putAt("Height", 1000)

        // freeze using col, row
        xFreeze.freezeAtPosition(2, 1)

    }

    static void insertEmployees(XSpreadsheet sht, List<Employee> el) {

        println("Inserting Employees")
        // column, row, string, sheet, flag
        // rows begin at 0, columns at 0
        int row = 1

        for (employee in el) {

            sht.setFormulaOfCell(0, row, employee.lastName)
            sht.setFormulaOfCell(1, row, employee.firstName)
            sht.setFormulaOfCell(2, row, employee.startDate.toString())
            sht.setFormulaOfCell(3, row, employee.email)
            row++
        }

    }


    static void main(String[] args) {

        // if we needed to use an instance
        // GroovyScript plainGroovy = new GroovyScript()

        List<Employee> employees

        mxRemoteContext = BootstrapSocketConnector.bootstrap(oooExeFolder)

        XComponentLoader aLoader = mxRemoteContext.componentLoader

        xComponent = aLoader.loadComponentFromURL(
                "private:factory/scalc", "_default", 0, new com.sun.star.beans.PropertyValue[0])

        XSpreadsheetDocument xSpreadsheetDocument = xComponent.getSpreadsheetDocument(mxRemoteContext)

        setupDocument(xSpreadsheetDocument)

        // get the spreadsheets from document
        System.out.println("Getting spreadsheet")
        XSpreadsheets xSheets = xSpreadsheetDocument.sheets

        XSpreadsheet xSpreadsheet = xSpreadsheetDocument.getSheetByName("Sheet1")

        employees = getEmployees("employees.csv")

        employees.each { employee ->
            println employee
        }

        formatSheet(xSpreadsheet, employees.size())

        insertEmployees(xSpreadsheet, employees)

        System.exit(0)

    }

}
