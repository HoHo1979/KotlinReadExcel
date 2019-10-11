
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId

val headers = mutableMapOf<Int,String>()
lateinit var localDate:LocalDate
var items= mutableListOf<Item>()
var location=""
val LOCATION="location"

enum class HEADERS(var value:String){
    NAME("name"),
    PACKAGE("package"),
    NUMBER_UNIT("number/unit"),
    UNIT("unit"),
    STOCK_DATE("stock/date"),
    LOWEST("lowest"),
    NOTE("note")
}

fun main()  {



    val wineExcel=FileInputStream(File("howinestock.xlsx"))
    val workbook=XSSFWorkbook(wineExcel)
    val sheet = workbook.getSheetAt(0)

    readHeaders(sheet)

    readDate(sheet)


    for (row in sheet) {


        var item = Item()

        loop@ for (header in headers) {
            var cell=row.getCell(header.key)

            when (cell?.cellType){
              CellType.NUMERIC->
                  when(header.value){
                      HEADERS.NUMBER_UNIT.value->item.number=cell.numericCellValue
                      HEADERS.LOWEST.value->item.lowestQuantity=cell.numericCellValue
                      HEADERS.STOCK_DATE.value->item.date=localDate
                  }
              CellType.STRING ->
                  when(header.value){
                      HEADERS.NAME.value-> {
                          //if cell has string value of location, read the next cell and break out of the loop
                          if(cell.stringCellValue.equals(LOCATION)){
                              var index=cell.columnIndex.inc()
                              cell = row.getCell(index)
                              if(cell.cellType==CellType.STRING){
                                  location = cell.stringCellValue
                                  break@loop
                              }
                          }

                          var itemName = cell.stringCellValue

                          item.name=itemName.split("/").get(0).trim()
                      }
                      HEADERS.PACKAGE.value->item._package=cell.stringCellValue
                      HEADERS.UNIT.value->item.unit=cell.stringCellValue
                      HEADERS.NOTE.value->item.note=cell.stringCellValue

                  }


            }

            item.location=location
        }
        //If item has name is empty string , don't add to the item list
        if(item.name!=""&&item.name!=HEADERS.NAME.value) {
            items.add(item)
        }
    }

    items.forEach { i-> println(i) }

    workbook.close()
    wineExcel.close()
}

private fun readDate(sheet: XSSFSheet) {
    val readDateRow = sheet.getRow(1)
    for (cell in readDateRow) {
        if (cell.cellType == CellType.NUMERIC) {
            val date = cell.dateCellValue
            localDate = LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault()).toLocalDate()
        }
    }
}

private fun readHeaders(sheet: XSSFSheet) {

    val headerRow = sheet.getRow(0)

    for (cell in headerRow) {
        if (cell.cellType == CellType.STRING) {
            headers.put(cell.columnIndex,cell.stringCellValue)
        }
    }
}
