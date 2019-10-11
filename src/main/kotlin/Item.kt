import java.time.LocalDate
import java.time.LocalDateTime

data class Item(var name:String, var _package:String, var number:Double, var unit:String, var date:LocalDate, var lowestQuantity:Double, var note:String,var location:String){

    constructor():this("","",0.0,"",LocalDate.now(),0.0,"","")

}
