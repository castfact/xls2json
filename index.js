// index.js

const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');

const app = express();
const port = 3000;

// Configure Multer for file upload
const storage = multer.memoryStorage();
const upload = multer({ storage });

// Serve the HTML file
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});

// Handle file upload and conversion
app.post('/upload', upload.single('file'), (req, res) => {
  // Load the uploaded Excel file from memory
  const workbook = new ExcelJS.Workbook();
  workbook.xlsx.load(req.file.buffer).then(() => {
    // Assume the first sheet is the one to convert
    const worksheet = workbook.getWorksheet(1);
    const jsonData = [];

    // Convert each row of the worksheet to JSON
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Exclude header row
        const rowData = {};
        
        row.eachCell((cell, colNumber) => {
          //console.log(worksheet.getRow(1).getCell(colNumber).value);    
          colName = worksheet.getRow(1).getCell(colNumber).value
          //rowData[`column${colNumber}`] = cell.value;
          if (worksheet.getRow(1).getCell(colNumber).value ==  "Ref") {
            
          } else if (worksheet.getRow(1).getCell(colNumber).value ==  "Skills" || 
              worksheet.getRow(1).getCell(colNumber).value ==  "Occupation" ||
              worksheet.getRow(1).getCell(colNumber).value ==  "Languages/ Dialect" || 
              worksheet.getRow(1).getCell(colNumber).value ==  "Alma Mater" || 
              worksheet.getRow(1).getCell(colNumber).value ==  "Nationality" ||
              worksheet.getRow(1).getCell(colNumber).value ==  "Tranport & Motoring License" ||
              worksheet.getRow(1).getCell(colNumber).value ==  "Body Type" ||
              worksheet.getRow(1).getCell(colNumber).value ==  "Eyes" ||
              worksheet.getRow(1).getCell(colNumber).value ==  "Hair" ||
              worksheet.getRow(1).getCell(colNumber).value ==  "Working Base" ) {
              val = cell.value.replaceAll("\n", ", ")
              arr = val.split(",");
            //loop through the array and trim each element

            var dict = {Skills : "skills", Occupation: "occupations", "Languages/ Dialect" : "languages", "Alma Mater" : "alma_maters", "Nationality" : "nationalities",
            "Tranport & Motoring License" : "licenses", "Body Type" : "body_types", "Eyes" : "eyes_colors", "Hair" : "hair_colors", "Working Base" : "working_bases"}

            //console.log("colName: " + dict[colName]);
            const mappedName = dict[colName]
            rowData[mappedName] = arr.map(
              function(item) { 
                return {"text" : item.trim()};
            })
          } else if(worksheet.getRow(1).getCell(colNumber).value ==  "Phone" || 
          worksheet.getRow(1).getCell(colNumber).value ==  "Email" ||
          worksheet.getRow(1).getCell(colNumber).value ==  "Address") {

            if (rowData["contacts"] == null) {
              rowData["contacts"] = [];
            }

            var val = cell.value.replaceAll("；", ";");
            val = val.trim();
            var arr = val.split("\n");
            //loop through the array and trim each element
            //console.log(worksheet.getRow(1).getCell(colNumber).value + ": " + arr.length + " " + val)
            if (arr.length > 0){
              var final_arr = arr.map(
                function(item) { 
                  var sub_arr = item.toString().split(";");
                  var keys = ["type", "category", "text"]
                  var dict = {};
  
                  //loop thru the arry and return the dict
                  dict["type"] = worksheet.getRow(1).getCell(colNumber).value
                  for(i = 1; i < keys.length; i++) {
                    if (sub_arr.length > (i)){
                      dict[keys[i]] = sub_arr[i - 1].trim()
                    } else {
                      dict[keys[i]] = ""
                    }
                  }
  
                  return dict;
              })
              //var arr = rowData["contacts"].values()
              for(i = 0; i < final_arr.length; i++) {
                rowData["contacts"].push(final_arr[i])
              }
              
            }
          }
          else if (worksheet.getRow(1).getCell(colNumber).value ==  "Movie" || 
            worksheet.getRow(1).getCell(colNumber).value ==  "TV Shows" ||
            worksheet.getRow(1).getCell(colNumber).value ==  "Commercials" || 
            worksheet.getRow(1).getCell(colNumber).value ==  "Variety Shows" || 
            worksheet.getRow(1).getCell(colNumber).value ==  "Performing Arts" ||
            worksheet.getRow(1).getCell(colNumber).value ==  "Broadcasts" ||
            worksheet.getRow(1).getCell(colNumber).value ==  "Modelling" ||
            worksheet.getRow(1).getCell(colNumber).value ==  "Voiceover" ||
            worksheet.getRow(1).getCell(colNumber).value ==  "Online & Multimedia" ||
            worksheet.getRow(1).getCell(colNumber).value ==  "Event" ) {
            
             var dict = {"Movie":"movies", "TV Shows":"tv_shows", "Commercials":"commercials", "Variety Shows":"variety_shows", "Performing Arts":"performing_arts", 
             "Broadcasts":"broadcasts", "Modelling":"modellings", "Voiceover":"voiceovers", "Online & Multimedia":"onlines", "Event":"events"}
  
            console.log("colName: " + dict[colName]);
            const mappedName = dict[colName]

            var val = cell.value.replaceAll("；", ";");
            val = val.trim();
            var arr = val.split("\n");
            //loop through the array and trim each element
            //console.log(worksheet.getRow(1).getCell(colNumber).value + ": " + arr.length + " " + val)
            rowData[mappedName] = (arr.length == 0) ? { } : arr.map(
              function(item) { 
                var sub_arr = item.toString().split(";");
                var keys = ["year", "name", "role_title", "role_name"]
                var dict = {};

                //loop thru the arry and return the dict
                for(i = 0; i < keys.length; i++) {
                  if (sub_arr.length > (i)){
                    dict[keys[i]] = sub_arr[i].trim()
                  } else {
                    dict[keys[i]] = ""
                  }
                }

                return dict;
            })
          } 
          else if (worksheet.getRow(1).getCell(colNumber).value ==  "Awards") {
            //award   
            var cellValue = (cell.value.trim() == "-") ? "" : cell.value
            console.log("cellvalue: " + cellValue)
            var val = cellValue.replaceAll("；", ";");
            val = val.trim();
            var arr = val.split("\n");
            //loop through the array and trim each element
            //console.log(worksheet.getRow(1).getCell(colNumber).value + ": " + arr.length + " " + val)
            rowData["awards"] = (arr.length == 0) ? { } : arr.map(
              function(item) { 
                var sub_arr = item.toString().split(";");
                var keys = ["year", "award_ceremony_name", "award_name", "winner"]
                var dict = {};

                //loop thru the arry and return the dict
                for(i = 0; i < keys.length; i++) {
                  if (sub_arr.length > (i)){
                    dict[keys[i]] = sub_arr[i].trim()
                  } else {
                    dict[keys[i]] = ""
                  }
                }

                return dict;
            })
          }
          else if (worksheet.getRow(1).getCell(colNumber).value ==  "Stage shows" || worksheet.getRow(1).getCell(colNumber).value ==  "Music Videos") {
            //music 
                var val = cell.value.replaceAll("；", ";");
                val = val.trim();
                var arr = val.split("\n");
                //loop through the array and trim each element
                //console.log(worksheet.getRow(1).getCell(colNumber).value + ": " + arr.length + " " + val)

              var dict = {"Stage shows":"stage_shows", "Music Videos":"music_videos"}
     
               console.log("colName: " + dict[colName]);
               const mappedName = dict[colName]
                   

                rowData[mappedName] = (arr.length == 0) ? { } : arr.map(
                  function(item) { 
                    var sub_arr = item.toString().split(";");
                    var keys = ["year", "name", "singer", "role_title", "role_name"]
                    var dict = {};
    
                    //loop thru the arry and return the dict
                    for(i = 0; i < keys.length; i++) {
                      if (sub_arr.length > (i)){
                        dict[keys[i]] = sub_arr[i].trim()
                      } else {
                        dict[keys[i]] = ""
                      }
                    }
    
                    return dict;
                })
          } else if (worksheet.getRow(1).getCell(colNumber).value ==  "Social Media") {
            //social media
            //agent
            var val = cell.value.replaceAll("@", "@");
                val = val.trim();
                var arr = val.split("\n");
                //loop through the array and trim each element
                //console.log(worksheet.getRow(1).getCell(colNumber).value + ": " + arr.length + " " + val)
                
                rowData["social_medias"] = (arr.length == 0) ? { } : arr.map(
                  function(item) { 
                    console.log(item.toString())
                    var sub_arr = item.toString().split("@");
                    console.log(sub_arr)
                    var keys = ["type", "category", "text"]
                    var dict = {};
    
                    //loop thru the arry and return the dict
                    for(i = 0; i < keys.length; i++) {
                      if (sub_arr.length > (i)){
                        dict[keys[i]] = sub_arr[i].trim()
                      } else {
                        dict[keys[i]] = ""
                      }
                    }
    
                    return dict;
                })
          } else if (worksheet.getRow(1).getCell(colNumber).value ==  "Agent/MGR") {
            //agent
            var val = cell.value.replaceAll("；", ";");
                val = val.trim();
                var arr = val.split("\n");
                //loop through the array and trim each element
                //console.log(worksheet.getRow(1).getCell(colNumber).value + ": " + arr.length + " " + val)
                rowData["agents"] = (arr.length == 0) ? { } : arr.map(
                  function(item) { 
                    var sub_arr = item.toString().split(";");
                    var keys = ["name", "phone", "email", "agent_status"]
                    var dict = {};
    
                    //loop thru the arry and return the dict
                    for(i = 0; i < keys.length; i++) {
                      if (sub_arr.length > (i)){
                        dict[keys[i]] = sub_arr[i].trim()
                      } else {
                        dict[keys[i]] = ""
                      }
                    }
    
                    return dict;
                })
          }  else if (worksheet.getRow(1).getCell(colNumber).value ==  "Date of Birth (DD/MM/YYYY)") {
            

            var date = new Date(cell.value);
            var value =  ((date.getDate() > 9) ? date.getDate() : ('0' + date.getDate())) + '/' + ((date.getMonth() > 8) ? (date.getMonth() + 1) : ('0' + (date.getMonth() + 1))) + '/' + date.getFullYear();
            
            //console.log("date datatype: " + typeof(cell.value) + " " + value)

            rowData["date_of_birth"] = value
          } else if (worksheet.getRow(1).getCell(colNumber).value ==  "Right-handed") {
            if (cell.value == "Y") {
              rowData["handedness"] = "Right handed"
            }
          } else if (worksheet.getRow(1).getCell(colNumber).value ==  "left-handed") {
            if (cell.value == "Y") {
              rowData["handedness"] = "Left handed"
            }
          } else {
            var dict = {"Gender" : "gender", "Age" : "age", "Years Active":"years_active", 
            "Height" : "height", "Weight" : "weight", "Skin Color" : "skin_color", "Right-handed" : "right_handed", "left-handed" : "left_handed", 
            "firstname_en" : "firstname_en", "lastname_en" : "lastname_en", "lastname_zh" : "lastname_zh", "firstname_zh" : "firstname_zh", "nickname" : "nickname",
            "Dress Size" : "dress_size",	"Shirt" : "shirt",	"Shoe" : "shoe", 	"Suit Coat" : "suit_cost_size",	"Pants" : "pants_size",	"Hat" : "hat_size", "Glove" : "glove"
            }
     
            

            console.log(colName + " / mapped name: " + dict[colName]);
            const mappedName = dict[colName]

            rowData[mappedName] = (cell.value == "-") ? "" : cell.value;
          }
        });
        rowData["name_display_format"] = ""
        jsonData.push(rowData);
      }
    });

    res.json(jsonData);
  });
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});