@Grab('com.opencsv:opencsv:5.2')

import com.opencsv.bean.CsvBindByName
import com.opencsv.bean.CsvDate
import groovy.transform.ToString

@ToString
class Employee {

    @CsvBindByName
    String lastName

    @CsvBindByName
    String firstName

    @CsvBindByName
    @CsvDate(value = "yyyy-MM-dd", writeFormatEqualsReadFormat = true)
    Date startDate

    @CsvBindByName
    String email

    }

