function myFunction() {
  var sheets = SpreadsheetApp.openById('')
  var source = sheets.getSheetByName('Source')
  var chart = sheets.getSheetByName('Chart')
//   Logger.log(source.getRange(2, 4).getValue())

  chart.getRange(3, 2, 24, 7).clear()
  chart.getRange(33, 2, 24, 7).clear()
  chart.getRange(63, 2, 24, 7).clear()
  chart.getRange(93, 2, 24, 7).clear()
   var hours;
   var myWeek = Utilities.formatDate(new Date(), "GMT", "w")
//   chart.getRange(1, 1).setValue(myWeek)
   var zone;
   var sourceZone;
   var offsetter = 30
   var shifter = 30
   var modifier = 0;
   var lesserMod = 0;
//   Logger.log(parseInt(zone))
   var values = source.getRange("A1:A").getValues();
   var lastline = values.filter(String).length;
   var myZone = parseInt(source.getRange(lastline, 3).getValue())
   var sheetOffset = 3.0
    
   if(chart.getRange(2, 1).isBlank() === true || chart.getRange(2, 1).getValue() == 0){
      chart.getRange(2, 1).setValue(Utilities.formatDate(new Date(), "GMT", "w"))
    }
    
   funcCurTimezone(myZone, chart)
   
   for(i=2;i<=lastline;i++){
     var offseti = i + 1 
     sourceZone = parseInt(source.getRange(i, 3).getValue())
     zone =  myZone - sourceZone
     
     var myDate = new Date(source.getRange(i, 4).getValue())
     var calcWeek = Utilities.formatDate(myDate, "GMT", "w")
     var latestWeek = parseInt(chart.getRange(2, 1).getValue())
//     Logger.log(calcWeek)
     var deltaWeek = parseInt(latestWeek) - parseInt(calcWeek)
     var datedWeek = parseInt(calcWeek) - parseInt(myWeek)
     var recovering = parseInt(myWeek) - parseInt(latestWeek)
     
     if(deltaWeek<-1 && recovering > 1){
       if(chart.getRange(2, 9).getValue() === ""){
         var picked = Date(chart.getRange(offseti, 4).getValue())
         weekDate(picked, chart)
       }
//       Logger.log("Insert One")
       var temp = insertRow(chart)
       chart.getRange(2, 1).setValue(parseInt(myWeek))
     }
     
     if(deltaWeek<0 && deltaWeek >=-1){
       if(chart.getRange(2, 9).getValue() === ""){
         var picked = Date(chart.getRange(offseti, 9).getValue())
         weekDate(picked, chart)
       }
//       Logger.log("Insert Two")
       var temp = insertRow(chart)
     }
     
      modifier = parseInt(matchWeek(deltaWeek))
      lesserMod = lesser(modifier)
      if(deltaWeek <=2 && datedWeek >= -1 ){
      
       for(j=5;j<=11;j++){
       
    
        hours = source.getRange(i, j).getValue().toString()
       if(hours){
           
           if(hours.indexOf(',')>-1){
             var arr = hours.split(",")
             
              for(r=0;r<=arr.length-1;r++){
//                Logger.log('wtf')
//                Logger.log(j.toString() + " " + i.toString())
                var rectified = parseInt(arr[r]) + zone
                if(rectified+3>2 && rectified <= 23 ){
                    chart.getRange(rectified+3 + shifter*modifier, j-sheetOffset).setNumberFormat("0")
                    chart.getRange(rectified+3 + shifter*modifier, j-sheetOffset).setValue(rectified)
                    chart.getRange(rectified+3 + shifter*modifier, j-sheetOffset).setBackgroundRGB(150, 150, 200)
                 }else if(rectified> 23){
                   
                   if(j==11){
                   latestWeek = parseInt(chart.getRange(2, 1).getValue())
                   deltaWeek = parseInt(latestWeek) - parseInt(calcWeek)
                    modifier = matchWeek(deltaWeek)
                     if(rectified + shifter*modifier < 30){
                     
                     if(chart.getRange(2, 9).getValue() === ""){
                           var picked = Date(chart.getRange(offseti, 9).getValue())
                           weekDate(picked, chart)
                       }
//                       Logger.log("Insert Three")
                       var temp = insertRow(chart)
                       
                       }
                       
//                       shifter += 30
                       chart.getRange(rectified+3 -24  + shifter*lesserMod, j - 6 -sheetOffset).setNumberFormat("0")
                       chart.getRange(rectified+3 -24  + shifter*lesserMod, j - 6 -sheetOffset).setValue(rectified - 24)
                       chart.getRange(rectified+3 -24  + shifter*lesserMod, j - 6 -sheetOffset).setBackgroundRGB(150, 150, 200)
                     
                   }else{
                     chart.getRange(rectified+3 -24 + shifter*modifier, j + 1 -sheetOffset).setNumberFormat("0")
                     chart.getRange(rectified+3 -24 + shifter*modifier, j + 1 -sheetOffset).setValue(rectified - 24)
                     chart.getRange(rectified+3 -24 + shifter*modifier, j + 1 -sheetOffset).setBackgroundRGB(150, 150, 200)
                   }
                 }else if(rectified + 3 <= 2){
                   if(j==5){
                   
                    chart.getRange(rectified+3 +24 + offsetter + shifter*modifier, j + 6 -sheetOffset).setNumberFormat("0")
                     chart.getRange(rectified+3 +24 + offsetter + shifter*modifier, j + 6 -sheetOffset).setValue(rectified + 24)
                     chart.getRange(rectified+3 +24 + offsetter + shifter*modifier, j + 6 -sheetOffset).setBackgroundRGB(150, 150, 200)

                   }else{
                     chart.getRange(rectified+3 +24 + shifter*modifier, j - 1 -sheetOffset).setNumberFormat("0")
                     chart.getRange(rectified+3 +24 + shifter*modifier, j - 1 -sheetOffset).setValue(rectified + 24)
                     chart.getRange(rectified+3 +24 + shifter*modifier, j - 1 -sheetOffset).setBackgroundRGB(150, 150, 200)
                   }
                 }
                
//                Logger.log(rectified)
//                Logger.log(chart.getRange(parseInt(arr[r])+2, j).getValue())
              }
           }else if(hours.indexOf(',')== -1 && hours !== "" ){
             
              var rectified = parseInt(hours) + zone
              if(rectified+3>2 && rectified <= 23){
                chart.getRange(rectified+3 + shifter*modifier, j-sheetOffset).setNumberFormat("0")
                chart.getRange(rectified+3 + shifter*modifier, j-sheetOffset).setValue(rectified)
                chart.getRange(rectified+3 + shifter*modifier, j-sheetOffset).setBackgroundRGB(150, 150, 200)
              }else if(rectified> 23){
                
                if(j==11){
                  latestWeek = parseInt(chart.getRange(2, 1).getValue())
                  deltaWeek = parseInt(latestWeek) - parseInt(calcWeek)
                  modifier = matchWeek(deltaWeek)
                   if(rectified + shifter*modifier < 30){
                       if(chart.getRange(2, 9).getValue() === ""){
                           var picked = Date(chart.getRange(i, 9).getValue())
                           weekDate(picked, chart)
                       }
                       
//                       Logger.log("Insert Four")
                       var temp = insertRow(chart)
                       
                       }
                       
                       
                       chart.getRange(rectified+3 -24 + shifter*lesserMod, j - 6 -sheetOffset).setNumberFormat("0")
                       chart.getRange(rectified+3 -24 + shifter*lesserMod , j - 6 -sheetOffset).setValue(rectified - 24)
                       chart.getRange(rectified+3 -24 + shifter*lesserMod , j - 6 -sheetOffset).setBackgroundRGB(150, 150, 200)
                       
                }else{
                  chart.getRange(rectified+3 -24 + shifter*modifier, j + 1 -sheetOffset).setNumberFormat("0")
                  chart.getRange(rectified+3 -24 + shifter*modifier, j + 1 -sheetOffset).setValue(rectified - 24)
                  chart.getRange(rectified+3 -24 + shifter*modifier, j + 1 -sheetOffset).setBackgroundRGB(150, 150, 200)
                }
              }else if(rectified + 3 <= 2){
                   if(offsetJ==5){
                     
                    chart.getRange(rectified+3 +24 + offsetter + shifter*modifier, j + 6 -sheetOffset).setNumberFormat("0")
                    chart.getRange(rectified+3 +24 + offsetter + shifter*modifier, j + 6 -sheetOffset).setValue(rectified + 24)
                    chart.getRange(rectified+3 +24 + offsetter + shifter*modifier, j + 6 -sheetOffset).setBackgroundRGB(150, 150, 200)
                     
                   }else{
                     chart.getRange(rectified+3 +24 + shifter*modifier, j - 1 -sheetOffset).setNumberFormat("0")
                     chart.getRange(rectified+3 +24 + shifter*modifier, j - 1 -sheetOffset).setValue(rectified + 24)
                     chart.getRange(rectified+3 +24 + shifter*modifier, j - 1 -sheetOffset).setBackgroundRGB(150, 150, 200)
                   }
                 }
              
//             Logger.log(rectified)
//             Logger.log(chart.getRange(parseInt(hours)+2, j).getValue())
           }
            
              
        }
      
     }
      }
      
     
    
   }
   
   drawSchedule(chart)
}

function insertRow(chart){
  chart.insertRowsBefore(2,30)
  var weekDays = chart.getRange(32, 2, 1, 7)
  weekDays.copyTo(chart.getRange(2, 2, 1, 7), {contentsOnly:true})
  
  var dayTime = chart.getRange(33, 1, 24, 1)
  dayTime.copyTo(chart.getRange(3, 1, 24, 1), {contentsOnly:true})
  
  var oldWeek = parseInt(chart.getRange(32, 1).getValue())
  var newWeek = oldWeek + 1
  chart.getRange(2, 1).setValue(parseInt(newWeek))
//  Logger.log(chart.getRange(2, 1).getValue())
  return 30
}


function matchWeek(deltaWeek){
  if(deltaWeek>=0){
    return deltaWeek
  }else{
    return 0
  }
  
   
  

}


function lesser(modifier){
  
  var a = parseInt(modifier) - 1
  
  if(a<0){
    return 0
  }
  
  return a

}



function weekDate(picked,chart){
  

//  var sheets = SpreadsheetApp.openById('157BtD_gU-_BHhaV9YCOir-K7dEXx6-0QPpIJB4xOlNI')
// 
//  var chart = sheets.getSheetByName('Chart')
  var startDate = new Date(picked)
  var endDate = new Date(picked)

  var inputWeek = parseInt(chart.getRange(2, 1).getValue())
  var checkWeek = parseInt(Utilities.formatDate(startDate, "GMT", "w"))
   
  if(inputWeek>1){
      while(checkWeek >= inputWeek){
      startDate = new Date(startDate.getTime() - 24*3600*1000)
      checkWeek = parseInt(Utilities.formatDate(startDate, "GMT", "w"))
    }

  }else if(inputWeek == 1){
  
      while(checkWeek > inputWeek){
      startDate = new Date(startDate.getTime() - 24*3600*1000)
      checkWeek = parseInt(Utilities.formatDate(startDate, "GMT", "w"))
    }
  }else{
  
       chart.getRange(2, 1).setValue(Utilities.formatDate(new Date(), "GMT", "w"))
  
  }

  
  if(chart.getRange(2, 9).isBlank()){
      startDate = new Date(startDate.getTime() + 24*3600*1000)
      chart.getRange(2, 9).setValue(startDate)
      
      endDate = new Date(startDate.getTime() + 6*24*3600*1000)
      chart.getRange(2, 10).setValue(endDate)
  }
  
  
}

function weekDateEX(calcDate,chart, row){
  

//  var sheets = SpreadsheetApp.openById('157BtD_gU-_BHhaV9YCOir-K7dEXx6-0QPpIJB4xOlNI')
// 
//  var chart = sheets.getSheetByName('Chart')
  var startDate = new Date(calcDate)
  var endDate = new Date(calcDate)

  var inputWeek = parseInt(chart.getRange(row, 1).getValue())
  var checkWeek = parseInt(Utilities.formatDate(startDate, "GMT", "w"))
   
  if(inputWeek>1){
      while(checkWeek >= inputWeek){
      startDate = new Date(startDate.getTime() - 24*3600*1000)
      checkWeek = parseInt(Utilities.formatDate(startDate, "GMT", "w"))
    }

  }else if(inputWeek == 1){
  
      while(checkWeek > inputWeek){
      startDate = new Date(startDate.getTime() - 24*3600*1000)
      checkWeek = parseInt(Utilities.formatDate(startDate, "GMT", "w"))
    }
  }else{
  
       chart.getRange(2, 1).setValue(Utilities.formatDate(new Date(), "GMT", "w"))
  
  }

  
  if(chart.getRange(row, 9).isBlank()){
      startDate = new Date(startDate.getTime())
      chart.getRange(row, 9).setValue(startDate)
      
      endDate = new Date(startDate.getTime() + 6*24*3600*1000)
      chart.getRange(row, 10).setValue(endDate)
  }
  
  
}
function funcCurTimezone(myZone,chart){
//  var sheets = SpreadsheetApp.getActive()
//  var chart = sheets.getSheetByName('Chart')
  var signedZone
  if(parseInt(myZone) >= 0){
    
    signedZone = "+" + myZone.toString()
  
  }else{
    signedZone = myZone.toString()
  }
  var showTimezone = "Current Timezone: GMT/UTC " + signedZone + "     (You can re-submit an empty Form with the wanted Timezone to view the schedule in your Timezone.)"
  chart.getRange(1, 1).setValue(showTimezone)
  
  
}



function drawSchedule(chart){
    var offsetter = 30.0
    var weekCount = chart.getRange(2, 1).getValue()
    Logger.log(chart.getRange(2, 1).getValue())
    var weekdays = chart.getRange(2, 2, 1, 7).getValues()
    var hours = chart.getRange(3, 1, 24, 1).getValues()
   
   
    for(i=0;i<=4;i++){
       
       chart.getRange(2 + offsetter*i, 2, 1, 7).setValues(weekdays)
       chart.getRange(3 + offsetter*i, 1, 24,1).setValues(hours)
       
        var row = 2.0+i*offsetter
        var rectWeekCount = chart.getRange(row, 1).getValue()
        var calcDate = getDateOfWeek(rectWeekCount,2020,chart.getRange(row, 1))
        
       
       
       if(i != 0 && weekCount-i > 0){
         chart.getRange(2 + offsetter*i, 1).setValue(weekCount - i)
       }
    
      weekDateEX(calcDate,chart,row)
    }
    
    var firstRow = chart.getRange("A:A")
    firstRow.setHorizontalAlignment("left")
    
    


}


function getDateOfWeek(w,y,row,chart){
  var d = 1 + (w - 1)*7
  Logger.log(Date(y,0,d))
  return new Date(y,0,d)
}




































