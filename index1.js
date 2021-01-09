const express = require('express');
const app = express();
const bodyParser = require('body-parser')
const fs = require('fs');
const fileUpload = require('express-fileupload')
const XlsxPopulate = require('xlsx-populate'); //used for formatting worksheet
const pino = require('pino');
const expressPino = require('express-pino-logger');
const {
   promisify
} = require("util");

const port = process.env.PORT || 3000;

var timeout = require('connect-timeout')

app.use(timeout(120000));

const logger = pino({
   prettyPrint: true
 });
const expressLogger = expressPino({ logger });

app.use(bodyParser.json({
   limit: '100mb'
}));

app.use(expressLogger);

app.use(fileUpload({
   createParentPath: true
}))

app.use((req, res, next) => {
   res.setHeader("Content-Type", "application/json");
   res.append('Access-Control-Allow-Origin', ['*']);
   res.append('Access-Control-Allow-Headers', 'Content-Type,Authorization');
   next()
});


app.listen(port, () => {
   logger.info('Server running on port %d', port);
})


app.post('/', async (req, res) => {

   if (req.files) {
      res.send({
         status: true,
         message: 'excel file uploaded'
      })
      const item = req.files;
      const file = item.file.name;
      const data = req.body
      const sheetObject = data.WriteToWorksheet
      const sheetName = data.Sheets

      logger.debug('excel file detected in request ==>', item.file.name)
      //console.log('json data detected ==>', data)

      for (const key of sheetObject) {
         const cellValue = key.WriteToCell; //header values
         const numValue = key.WriteNumberToCell;
         const cellBorder = key.SetCellBorders; //cell or column ranges for borders
         const bgColor = key.SetCellBackgroundColor; //background fill color used mostly for header range
         const columnWidth = key.SetColumnWidth; //width formats for columns
         const cellFont = key.SetCellFont; //font settings
         const numberFormat = key.SetCellNumberFormat //cell ranges for cells to be set as number format in worksheet
         const autoFilter = key.SetAutoFilter
         const formula = key.WriteFormulaToCell
         const worksheet = key.SelectWorksheet


         XlsxPopulate.fromFileAsync(file).then(workbook => {

            //console.log('sheets ===>',sheetName)
            const sheet = workbook.sheet(worksheet).name(sheetName.toString())
            //parse Headers into workbook
            const printHeaders = () => new Promise((resolve) => {
               logger.info('writing headers and text to excel')
               let pr = Promise.resolve(0);
               for (const data of cellValue) {
                  pr = pr.then((val) => {
                     let cellRef = data.Cell;
                     let cellValue = data.Value;
                     //const str = cellRef.replace(/([a-zA-Z0-9-]+):([a-zA-Z0-9-]+)/g, "\"$1\:\$2\"");
                     //console.log('cell values ===>', cellValue)
                     const r = workbook.sheet(worksheet).range(cellRef)
                     r.value([
                        cellValue,

                     ]);
                  })

               };
               resolve(pr)
            })
            //write numbers into workbook
            const printNumbers = () => new Promise((resolve) => {
               logger.info('writing numbers to excel')
               let pr = Promise.resolve(0);
               for (const num of numValue) {
                  pr = pr.then((val) => {
                     let cellRef = num.Cell;
                     let cellValue = num.Number;
                     //console.log('num values ====>', cellValue)
                     const r = workbook.sheet(worksheet).range(cellRef)
                     r.value(cellValue)
                  })

               };
               resolve(pr)
            })

            //To be invoked sequentially where printHeaders has to be resolved or rejected before printNumbers   
            async function foo() {
               try {
                  await printHeaders()
                  await printNumbers()
               } catch (err) {
                  res.status(600).send({
                     error: err.message
                  })
               }

            }

            //set column widths
            columnWidth.forEach(item => {
               //let sheet = workbook.sheet(worksheet)
               //console.log('column width cells ===> ', item.Cell)

               //const test = item.Cell
               for (let i = 0; i < columnWidth.length; i++) {
                  sheet.column(item.Cell).width(item.Width)
               }
            });


            //Formatting fonts
            cellFont.forEach(item => {
               let sheet = workbook.sheet(worksheet)
               //console.log('fonts ====>',item)
               let cellRange = item.Cell;
               let fontFamily = item.FontName;
               let fontSize = item.Size
               let fontStyle = item.Style
               const range = sheet.range(item.Cell)
               range.forEach(item => {
                  item.style({
                     fontFamily: "Verdana",
                     fontSize: fontSize,
                     bold: fontStyle
                  })
               })
            });

            //set column borders
            cellBorder.forEach(item => {
               let sheet = workbook.sheet(worksheet)
               const range = sheet.range(item.Cell)
               const borderRange = item.Border
               console.log('border range ===>', borderRange)
               range.forEach(item => {
                  if (borderRange == "Top") {
                     item.style({
                        topBorder: true
                     })
                  } else if (borderRange == "Right") {
                     item.style({
                        rightBorder: true
                     })
                  } else if (borderRange == "Left") {
                     item.style({
                        leftBorder: true
                     })
                  } else if (borderRange == "Bottom") {
                     item.style({
                        bottomBorder: true
                     })
                  }
               })


            });

            //set autofilter range
            autoFilter.forEach(item => {
               let sheet = workbook.sheet(worksheet)
               const range = sheet.range(item.Cell)
               //console.log('auto filter range: ', range)
               range.autoFilter()

            })

            //set formula range
            formula.forEach(item => {
               let sheet = workbook.sheet(worksheet)
               //console.log('formula item: ',item)
               const range = sheet.range(item.Cell)
               const formula = item.FormulaValue
               console.log('formula value: ', formula)

               range.formula(formula)
            })

            //Background fills for cell range
            bgColor.forEach(item => {
               let sheet = workbook.sheet(worksheet)
               const range = sheet.range(item.Cell)
               range.forEach(item => {
                  item.style("fill", {
                     type: "solid",
                     color: {
                        rgb: "E0E0E0"
                     }

                  })
               });
            });


            // Write to file and send base64 in response. 
            const output = () => new Promise((resolve) => {
               let pr = Promise.resolve(0);
               workbook.outputAsync("base64").then(function (base64) {
                  pr = pr.then((val) => {
                     res.send(base64)
                     workbook.toFileAsync(file)
                  })

               })
               resolve(pr)
            })

            async function final() {
               try {
                  await foo();
                  await output()
               } catch (err) {
                  res.status(701).send({
                     error: err.message
                  })
               }
            };

            final().catch(err => console.error(err))


         })

      }

   } else {
      logger.debug('no excel file detected in request')
      const data = req.body
      const excelFile = req.body.File

      //if base64 encoded file is detected in body, write data to base64 file
      if (excelFile != null) {
         logger.debug('base64 file detected')

         try {
            logger.debug('decoding base64 to file')
            fs.writeFileSync('test.xlsx', excelFile, 'base64')
         } catch (err) {
            res.status(501).send({
               error: err.message
            })
         }

         try {
            const data = req.body
            const sheetNames = data.Sheets
            //console.log('no of sheets ====>', sheetNames)

            if (sheetNames.length > 1) {
               console.log('More sheets <====')

               const sheetObject = data.WriteToWorksheet
               //console.log('sheytt==>>>>', sheetObject.length)

               const workbook = await XlsxPopulate.fromFileAsync("./test.xlsx");

               const allBase64 = await Promise.all(
                  sheetObject.map(async (key, sheetIndex) => {
                        try {
                           const cellValue = key.WriteToCell; //header values
                           const numValue = key.WriteNumberToCell;
                           const cellBorder = key.SetCellBorders; //cell or column ranges for borders
                           const bgColor = key.SetCellBackgroundColor; //background fill color used mostly for header range
                           const columnWidth = key.SetColumnWidth; //width formats for columns
                           const cellFont = key.SetCellFont; //font settings
                           const numberFormat = key.SetCellNumberFormat //cell ranges for cells to be set as number format in worksheet
                           const autoFilter = key.SetAutoFilter
                           const formula = key.WriteFormulaToCell
                           const worksheet = key.SelectWorksheet

                           const currentSheetName = sheetNames[sheetIndex];

                           const printHeaders = () => {
                              return Promise.all(cellValue.map(item => {
                                 try {
                                    let cellRef = item.Cell;
                                    let cellValue = item.Value
                                    //console.log('current sheet name ====>', currentSheetName)
                                    let newValue = cellValue.map(item => {
                                       return item.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, '')
                                       //return item.map(key => key.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, ''))
                                    })
                                    const r = workbook.sheet(worksheet).range(cellRef);
                                    return r.value([newValue])
                                 } catch (err) {
                                    try {
                                       let cellRef = item.Cell;
                                       let cellValue = item.Value
                                       let newValue = cellValue.map((item) => {
                                          return item.map(key => key.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, ''))
                                          //return item
                                       })

                                       const r = workbook.sheet(worksheet).cell(cellRef)
                                       return r.value([newValue])

                                    } catch (err) {
                                       res.send({
                                          err: err.message
                                       })
                                    }
                                 }
                              }));
                           }

                           //write numbers into workbook
                           const printNumbers = () => {
                              return Promise.all(

                                 numValue.map(num => {
                                    const cellRef = num.Cell;
                                    const cellValue = num.Number;
                                    //console.log('num values ====>', cellValue + ' ' + 'num cells ===>', cellRef + ' ' + worksheet)
                                    const r = workbook.sheet(currentSheetName).range(cellRef)
                                    return r.value(cellValue)
                                 })
                              )
                           }

                           //set column widths

                           const setWidth = async () => {
                              await Promise.all(columnWidth.map(item => {
                                 let sheet = workbook.sheet(currentSheetName)

                                 for (let i = 0; i < columnWidth.length; i++) {
                                    sheet.column(item.Cell).width(item.Width)
                                 }
                              }))
                           }



                           //set cell font
                           const setFont = async () => {
                              await Promise.all(cellFont.map(item => {
                                 let sheet = workbook.sheet(worksheet)
                                 let fontFamily = item.FontName;
                                 let fontSize = item.Size
                                 let fontStyle = item.Style
                                 const range = sheet.range(item.Cell)
                                 range.forEach(item => {
                                    item.style({
                                       fontFamily: "Verdana",
                                       fontSize: fontSize,
                                       bold: fontStyle
                                    })
                                 })
                              }));
                           }

                           try {
                              await printHeaders();
                              await printNumbers();
                              await setWidth();
                              await setFont();

                           } catch (err) {
                              try {
                                 await printHeaders();
                                 await printNumbers();
                                 //await setWidth();
                                 //await setFont();
                              } catch (err) {
                                 try {
                                    //if WriteteToCell is 2d array in req & no Width / Font object
                                    await printHeaders();
                                 } catch (err) {
                                    res.send({
                                       err: err.message
                                    })
                                 }

                              }
                           };

                           //set column borders
                           if (cellBorder) {
                              cellBorder.forEach(item => {
                                 let sheet = workbook.sheet(worksheet)
                                 const range = sheet.range(item.Cell)
                                 const borderRange = item.Border
                                 //console.log('border range ===>', borderRange)
                                 range.forEach(item => {
                                    if (borderRange == "Top") {
                                       item.style({
                                          topBorder: true
                                       })
                                    } else if (borderRange == "Right") {
                                       item.style({
                                          rightBorder: true
                                       })
                                    } else if (borderRange == "Left") {
                                       item.style({
                                          leftBorder: true
                                       })
                                    } else if (borderRange == "Bottom") {
                                       item.style({
                                          bottomBorder: true
                                       })
                                    }
                                 })
                              });
                           }

                           //
                           //set autofilter range
                           if (autoFilter) {
                              autoFilter.forEach(item => {
                                 let sheet = workbook.sheet(worksheet)
                                 const range = sheet.range(item.Cell)
                                 //console.log('auto filter range: ', range)
                                 range.autoFilter()

                              })
                           }

                           //
                           //set formula range
                           if (formula) {
                              formula.forEach(item => {
                                 let sheet = workbook.sheet(currentSheetName)
                                 const range = sheet.range(item.Cell)
                                 const formula = item.FormulaValue
                                 //console.log('formula value: ', formula)

                                 range.formula(formula)
                              })
                           }

                           //
                           //Background fills for cell range
                           if (bgColor) {
                              bgColor.forEach(item => {
                                 let sheet = workbook.sheet(currentSheetName)
                                 const range = sheet.range(item.Cell)
                                 range.forEach(item => {
                                    item.style("fill", {
                                       type: "solid",
                                       color: {
                                          rgb: "E0E0E0"
                                       }

                                    })
                                 });
                              });
                           }

                        } catch (err) {
                           res.status(701).send({
                              err: err.message
                           })
                        }
                     }


                  )
               )
               // Write to file and send base64 in response. 
               const base64 = () => new Promise((resolve) => {
                  let pr = Promise.resolve(0);
                  workbook.outputAsync("base64").then(function (base64) {
                     pr = pr.then((val) => {
                        res.send(base64)

                     })

                  })
                  resolve(pr)
               })


               await fs.unlinkSync('./test.xlsx');
               await base64();
               //const base64 = await workbook.outputAsync("base64")
               //await workbook.toFileAsync(`./tmp.xlsx`)
               //res.send(base64);


            } else if (sheetNames.length < 2) {

               logger.debug('Creating Single Sheet excel file')
               const sheetObject = data.WriteToWorksheet
               const sheetName = data.Sheets

               for (const key of sheetObject) {
                  //console.log(key)
                  const cellValue = key.WriteToCell; //header values
                  const numValue = key.WriteNumberToCell;
                  const cellBorder = key.SetCellBorders; //cell or column ranges for borders
                  const bgColor = key.SetCellBackgroundColor; //background fill color used mostly for header range
                  const columnWidth = key.SetColumnWidth; //width formats for columns
                  const cellFont = key.SetCellFont; //font settings
                  const numberFormat = key.SetCellNumberFormat //cell ranges for cells to be set as number format in worksheet
                  const autoFilter = key.SetAutoFilter
                  const formula = key.WriteFormulaToCell
                  const worksheet = key.SelectWorksheet



                  //Initialize Excel workbook
                  XlsxPopulate.fromFileAsync("./test.xlsx").then(workbook => {

                     //console.log('sheets ===>',sheetName)
                     const sheet = workbook.sheet(worksheet).name(sheetName.toString())
                     //parse Headers into workbook
                     const printHeaders = () => new Promise((resolve) => {

                        let pr = Promise.resolve(0);
                        for (const data of cellValue) {
                           pr = pr.then((val) => {
                              let cellRef = data.Cell;
                              let cellValue = data.Value;
                              let newValue = cellValue.map(item => {
                                 return item.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, '')
                                 //return item.map(key => key.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, ''))
                                 //return item
                              })
                              const r = workbook.sheet(worksheet).range(cellRef)
                              r.value([newValue]);
                           })

                        };
                        resolve(pr)

                     })
                     //write numbers into workbook
                     const printNumbers = () => {
                        return Promise.all(

                           numValue.map(num => {

                              const cellRef = num.Cell;
                              const cellValue = num.Number;
                              console.log('num values ====>', cellValue + ' ' + 'num cells ===>', cellRef + ' ' + worksheet)
                              const r = workbook.sheet(worksheet).range(cellRef)
                              return r.value(cellValue)
                           })
                        )
                     }

                     //To be invoked sequentially where printHeaders has to be resolved or rejected before printNumbers   
                     async function foo() {
                        try {
                           await printHeaders()
                           await printNumbers()
                        } catch (err) {
                           try {
                              await printHeaders()
                           } catch (err) {
                              res.status(600).send({
                                 error: err.message
                              })
                           }

                        }

                     }

                     //set column widths
                     if (columnWidth) {
                        columnWidth.forEach(item => {
                           //let sheet = workbook.sheet(worksheet)
                           //console.log('column width cells ===> ', item.Cell)

                           //const test = item.Cell
                           for (let i = 0; i < columnWidth.length; i++) {
                              sheet.column(item.Cell).width(item.Width)
                           }
                        });
                     }



                     //Formatting fonts
                     if (cellFont) {
                        cellFont.forEach(item => {
                           let sheet = workbook.sheet(worksheet)
                           //console.log('fonts ====>',item)
                           let cellRange = item.Cell;
                           let fontFamily = item.FontName;
                           let fontSize = item.Size
                           let fontStyle = item.Style
                           const range = sheet.range(item.Cell)
                           range.forEach(item => {
                              item.style({
                                 fontFamily: "Verdana",
                                 fontSize: fontSize,
                                 bold: fontStyle
                              })
                           })
                        });
                     }


                     //set column borders
                     if (cellBorder) {
                        cellBorder.forEach(item => {
                           let sheet = workbook.sheet(worksheet)
                           const range = sheet.range(item.Cell)
                           const borderRange = item.Border
                           //console.log('border range ===>', borderRange)
                           range.forEach(item => {
                              if (borderRange == "Top") {
                                 item.style({
                                    topBorder: true
                                 })
                              } else if (borderRange == "Right") {
                                 item.style({
                                    rightBorder: true
                                 })
                              } else if (borderRange == "Left") {
                                 item.style({
                                    leftBorder: true
                                 })
                              } else if (borderRange == "Bottom") {
                                 item.style({
                                    bottomBorder: true
                                 })
                              }
                           })


                        });
                     }


                     //set autofilter range   
                     if (autoFilter) {
                        autoFilter.forEach(item => {
                           let sheet = workbook.sheet(worksheet)
                           const range = sheet.range(item.Cell)
                           //console.log('auto filter range: ', range)
                           range.autoFilter()

                        })
                     }


                     //set formula range
                     if (formula) {
                        formula.forEach(item => {
                           let sheet = workbook.sheet(worksheet)
                           //console.log('formula item: ',item)
                           const range = sheet.range(item.Cell)
                           const formula = item.FormulaValue
                           //console.log('formula value: ', formula)

                           range.formula(formula)
                        })
                     }


                     //Background fills for cell range
                     if (bgColor) {
                        bgColor.forEach(item => {
                           let sheet = workbook.sheet(worksheet)
                           const range = sheet.range(item.Cell)
                           range.forEach(item => {
                              item.style("fill", {
                                 type: "solid",
                                 color: {
                                    rgb: "E0E0E0"
                                 }

                              })
                           });
                        });
                     }



                     // Write to file and send base64 in response. 
                     const output = () => new Promise((resolve) => {
                        let pr = Promise.resolve(0);
                        workbook.outputAsync("base64").then(function (base64) {
                           pr = pr.then((val) => {
                              res.send(base64)
                              //workbook.toFileAsync("./tmp.xlsx")
                           })

                        })
                        resolve(pr)
                     })

                     async function final() {
                        try {
                           await foo();
                           await output()
                           await fs.unlinkSync('./test.xlsx');
                        } catch (err) {
                           res.status(701).send({
                              error: err.message
                           })
                        }
                     };

                     final()

                  })
               };

            }
         } catch (err) {
            console.log(err)
         }


      } else {
         //if req.body just contains json data, create new workbook
         logger.info('No excel file or base64 in request. Creating blank excel file')
         try {
            const data = req.body
            const sheetNames = data.Sheets
            //logger.debug('Multiple excel sheets created ==>', sheetNames)

            if (sheetNames.length > 1) {
              
               const sheetObject = data.WriteToWorksheet
               //console.log('sheytt==>>>>', sheetObject.length)

               const workbook = await XlsxPopulate.fromBlankAsync();

               const allBase64 = await Promise.all(
                  sheetObject.map(async (key, sheetIndex) => {
                        try {
                           const cellValue = key.WriteToCell; //header values
                           const numValue = key.WriteNumberToCell;
                           const cellBorder = key.SetCellBorders; //cell or column ranges for borders
                           const bgColor = key.SetCellBackgroundColor; //background fill color used mostly for header range
                           const columnWidth = key.SetColumnWidth; //width formats for columns
                           const cellFont = key.SetCellFont; //font settings
                           const numberFormat = key.SetCellNumberFormat //cell ranges for cells to be set as number format in worksheet
                           const autoFilter = key.SetAutoFilter
                           const formula = key.WriteFormulaToCell
                           const worksheet = key.SelectWorksheet

                           const currentSheetName = sheetNames[sheetIndex];
                           //console.log('current sheet ===>', currentSheetName)

                           const printSheets = async () => {
                              try {
                                 return workbook.addSheet(currentSheetName);
                              } catch (err) {
                                 logger.debug('errrr==>>>', err)
                              }
                           }

                           const printHeaders = () => {
                              return Promise.all(cellValue.map(item => {
                                 try {
                                    let cellRef = item.Cell;
                                    let cellValue = item.Value;
                                    //console.log('header values ====>', cellValue)
                                    let newValue = cellValue.map(item => {
                                       return item.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, '')
                                    });
                                    const r = workbook.sheet(currentSheetName).range(cellRef)
                                    return r.value([newValue])
                                 } catch (err) {
                                    res.send(err)
                                    logger.debug('error writing headers/text', err)
                                 }
                              }))
                           }

                           const print2dArray = () => {
                              return Promise.all(cellValue.map(item => {
                                 try {
                                    let cellRef = item.Cell;
                                    let cellValue = item.Value
                                    cellValue.forEach(item => {
                                       item.map((item) => {
                                          if (typeof item == 'string') {
                                             item.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, '')
                                          }
                                       })
                                    })
                                    // let newValue = cellValue.map((item) => {
                                    //    return item.map(key => key.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, ''))
                                    //    //return item
                                    // })

                                    const r = workbook.sheet(currentSheetName).cell(cellRef)
                                    return r.value([cellValue])
                                 } catch (err) {
                                    res.send(err)
                                    logger.debug('Error processing 2d array', err);
                                 }
                              }));
                           }

                           //write numbers into workbook
                           const printNumbers = () => {
                              return Promise.all(

                                 numValue.map(num => {
                                    const cellRef = num.Cell;
                                    const cellValue = num.Number;
                                    const newValue = JSON.parse(cellValue)
                                    //console.log('num values ====>', newValue)
                                    const r = workbook.sheet(currentSheetName).range(cellRef)
                                    return r.value(newValue)
                                 })
                              )
                           }


                           //set column widths
                           const setWidth = async () => {
                              await Promise.all(columnWidth.map(item => {
                                 let sheet = workbook.sheet(currentSheetName)

                                 for (let i = 0; i < columnWidth.length; i++) {
                                    sheet.column(item.Cell).width(item.Width)
                                 }
                              }))
                           }

                           //set cell font
                           const setFont = async () => {
                              await Promise.all(cellFont.map(item => {
                                 let sheet = workbook.sheet(currentSheetName)
                                 let fontFamily = item.FontName;
                                 let fontSize = item.Size
                                 let fontStyle = item.Style
                                 const range = sheet.range(item.Cell)
                                 range.forEach(item => {
                                    item.style({
                                       fontFamily: "Verdana",
                                       fontSize: fontSize,
                                       bold: fontStyle
                                    })
                                 })
                              }));
                           }

                           //To be invoked sequentially where printHeaders has to be resolved or rejected before printNumbers                        
                           const printAndFormat = async () => {
                              try {
                                 await printSheets();
                                 await printHeaders();
                                 await printNumbers();
                                 //await setWidth();
                                 //await setFont();
                              } catch (err) {
                                 // if WriteNumberToCell object is not present in req
                                 try {
                                    await printSheets();
                                    await printHeaders();
                                    await setWidth();
                                    await setFont();
                                 } catch (err) {
                                    // if using 2d arrays from req
                                    try {
                                       await printSheets();
                                       await print2dArray();
                                       await setWidth();
                                       await setFont();
                                    } catch (err) {
                                       logger.debug(err)
                                       res.status(502).send({
                                          err: err.message                                          
                                       })
                                    }
                                 }

                              }
                           }
                           try {
                              await printAndFormat();
                           } catch (err) {
                              logger.info(err)
                              res.send(err)
                           }

                           //set column borders
                           if(cellBorder ) {
                              cellBorder.forEach(item => {
                                 let sheet = workbook.sheet(currentSheetName)
                                 const range = sheet.range(item.Cell)
                                 const borderRange = item.Border
                                 //console.log('border range ===>', borderRange)
                                 range.forEach(item => {
                                    if (borderRange == "Top") {
                                       item.style({
                                          topBorder: true
                                       })
                                    } else if (borderRange == "Right") {
                                       item.style({
                                          rightBorder: true
                                       })
                                    } else if (borderRange == "Left") {
                                       item.style({
                                          leftBorder: true
                                       })
                                    } else if (borderRange == "Bottom") {
                                       item.style({
                                          bottomBorder: true
                                       })
                                    }
                                 })
                              });
                           }
                           
                           //
                           //set autofilter range
                           if(autoFilter) {
                              autoFilter.forEach(item => {
                                 let sheet = workbook.sheet(currentSheetName)
                                 const range = sheet.range(item.Cell)
                                 //console.log('auto filter range: ', range)
                                 range.autoFilter()
   
                              })
                           }
                           
                           //
                           //set formula range
                           if (formula) {
                              formula.forEach(item => {
                                 let sheet = workbook.sheet(currentSheetName)
                                 const range = sheet.range(item.Cell)
                                 const formula = item.FormulaValue
                                 //console.log('formula value: ', formula)

                                 range.formula(formula)
                              })
                           }

                           //
                           //Background fills for cell range
                           if(bgColor) {
                              bgColor.forEach(item => {
                                 let sheet = workbook.sheet(currentSheetName)
                                 const range = sheet.range(item.Cell)
                                 range.forEach(item => {
                                    item.style("fill", {
                                       type: "solid",
                                       color: {
                                          rgb: "E0E0E0"
                                       }
   
                                    })
                                 });
                              });
                           };
        
                        } catch (err) {
                           res.send(err);
                           logger.info(err)
                        }
                     }


                  )
               )

               // Write to file and send base64 in response. 
               const base64 = () => new Promise((resolve) => {

                  let pr = Promise.resolve(0);
                  workbook.outputAsync("base64").then(function (base64) {
                     pr = pr.then((val) => {
                        res.send(base64)
                        logger.info('base64 response sent')
                     })

                  })
                  resolve(pr)
               })


               await workbook.deleteSheet('Sheet1');
               await base64();
               //const base64 = await workbook.outputAsync("base64")
               await workbook.toFileAsync(`./tmp.xlsx`)
               //res.send(base64);


            } else if (sheetNames.length < 2) {
               logger.debug('Single sheet detected ==>', sheetNames)
               logger.info('No excel file or base 64 in request. Creating blank excel file')
               const sheetObject = data.WriteToWorksheet
               const sheetName = data.Sheets

               for (const key of sheetObject) {
                  //console.log(key)
                  const cellValue = key.WriteToCell; //header values
                  const numValue = key.WriteNumberToCell;
                  const cellBorder = key.SetCellBorders; //cell or column ranges for borders
                  const bgColor = key.SetCellBackgroundColor; //background fill color used mostly for header range
                  const columnWidth = key.SetColumnWidth; //width formats for columns
                  const cellFont = key.SetCellFont; //font settings
                  const numberFormat = key.SetCellNumberFormat //cell ranges for cells to be set as number format in worksheet
                  const autoFilter = key.SetAutoFilter
                  const formula = key.WriteFormulaToCell
                  const worksheet = key.SelectWorksheet

                  //Initialize Excel workbook
                  XlsxPopulate.fromBlankAsync().then(workbook => {

                     //console.log('sheets ===>',sheetName
                     const sheet = workbook.sheet(worksheet).name(sheetName.toString())
                     logger.info('Excel worksheet initialized')
                     //parse Headers into workbook
                     const printHeaders = () => new Promise((resolve) => {
                        let pr = Promise.resolve(0);
                        for (const data of cellValue) {
                           pr = pr.then((val) => {
                              let cellRef = data.Cell;
                              let cellValue = data.Value;
                              let newValue = cellValue.map(item => {
                                 return item.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, '')
                              })
                              const r = workbook.sheet(worksheet).range(cellRef)
                              r.value([newValue]);

                           })

                        };
                        resolve(pr)
                     })
                     const print2dArray = () => new Promise((resolve) => {

                        let pr = Promise.resolve(0);
                        for (const data of cellValue) {
                           pr = pr.then((val) => {
                              let cellRef = data.Cell;
                              let cellValue = data.Value;
                              //console.log( typeof cellValue) 
                              cellValue.forEach(item => {

                                 item.map((item) => {
                                    if (typeof item == 'string') {
                                       //console.log('whats my string item ====>',item)
                                       item.replace(/[^A-Za-z 0-9 \.,\?""!@#\$%\^&\*\(\)-_=\+;:<>\/\\\|\}\{\[\]`~]*/g, '')
                                    }
                                 })

                              });

                              const r = workbook.sheet(worksheet).cell(cellRef)

                              r.value([
                                 [cellValue]
                              ])

                           })

                        };
                        resolve(pr)
                     })


                     //write numbers into workbook
                     const printNumbers = () => new Promise((resolve) => {
                        let pr = Promise.resolve(0);
                        for (const num of numValue) {
                           pr = pr.then((val) => {
                              let cellRef = num.Cell;
                              let cellValue = num.Number;
                              //console.log('num values ====>', cellValue)
                              const r = workbook.sheet(worksheet).range(cellRef)
                              r.value(cellValue)
                           })

                        };
                        resolve(pr)
                     })

                     //To be invoked sequentially where printHeaders has to be resolved or rejected before printNumbers   
                     async function foo() {
                        try {
                           await printHeaders()
                           await printNumbers();
                           logger.debug('data written to excel')
                        } catch (err) {
                           //WriteNumToCell may not be included in req in some cases
                           try {
                              await printHeaders();

                           } catch (err) {
                              try {
                                 await print2dArray();

                              } catch (err) {
                                 logger.info('Single Sheet blank workbook error ===>',err)
                                 res.status(600).send({
                                    error: err.message
                                 })
                              }
                           }
                        }
                     }


                     //set column widths
                     columnWidth.forEach(item => {
                        //let sheet = workbook.sheet(worksheet)
                        //console.log('column width cells ===> ', item.Cell)

                        //const test = item.Cell
                        for (let i = 0; i < columnWidth.length; i++) {
                           sheet.column(item.Cell).width(item.Width)
                        }
                     });


                     //Formatting fonts
                     cellFont.forEach(item => {
                        let sheet = workbook.sheet(worksheet)
                        //console.log('fonts ====>',item)
                        let cellRange = item.Cell;
                        let fontFamily = item.FontName;
                        let fontSize = item.Size
                        let fontStyle = item.Style
                        const range = sheet.range(item.Cell)
                        //console.log('what is my range ===>', cellRange)
                        range.forEach(item => {
                           item.style({
                              fontFamily: "Verdana",
                              fontSize: fontSize,
                              bold: fontStyle                              
                           })
                        })
                     });

                     //set column borders
                     cellBorder.forEach(item => {
                        let sheet = workbook.sheet(worksheet)
                        const range = sheet.range(item.Cell)
                        const borderRange = item.Border
                        //console.log('border range ===>', borderRange)
                        range.forEach(item => {
                           if (borderRange == "Top") {
                              item.style({
                                 topBorder: true
                              })
                           } else if (borderRange == "Right") {
                              item.style({
                                 rightBorder: true
                              })
                           } else if (borderRange == "Left") {
                              item.style({
                                 leftBorder: true
                              })
                           } else if (borderRange == "Bottom") {
                              item.style({
                                 bottomBorder: true
                              })
                           }
                        })


                     });

                     //set autofilter range
                     autoFilter.forEach(item => {
                        let sheet = workbook.sheet(worksheet)
                        const range = sheet.range(item.Cell)
                        //console.log('auto filter range: ', range)
                        range.autoFilter()

                     })

                     //set formula range
                     if (formula) {
                        formula.forEach(item => {
                           let sheet = workbook.sheet(worksheet)
                           //console.log('formula item: ',item)
                           const range = sheet.range(item.Cell)
                           const formula = item.FormulaValue
                           //console.log('formula value: ', formula)

                           range.formula(formula)
                        })

                     }

                     //Background fills for cell range
                     bgColor.forEach(item => {
                        let sheet = workbook.sheet(worksheet)
                        const range = sheet.range(item.Cell)
                        range.forEach(item => {
                           item.style("fill", {
                              type: "solid",
                              color: {
                                 rgb: "E0E0E0"
                              }

                           })
                        });
                     });


                     // Write to file and send base64 in response. 
                     const output = () => new Promise((resolve) => {
                        const values = workbook.sheet("Sheet1").usedRange().value();

                        let pr = Promise.resolve(0);
                        workbook.outputAsync("base64").then(function (base64) {
                           pr = pr.then((val) => {
                              res.send(base64)
                              workbook.toFileAsync("./tmp.xlsx")
                              logger.info('base64 response sent')
                           })

                        })
                        resolve(pr)
                     })

                     async function final() {
                        try {
                           await foo();
                           await output()
                        } catch (err) {
                           logger.debug(err)
                           res.status(701).send({
                              error: err.message
                           })
                        }
                     };

                     final()

                  })
               };

            }
         } catch (err) {
            logger.info(err)
         }


         //const writePath = data.SaveAs
         //WriteToWorksheet object in json file



      }

   }
})


//fs.readFile('./text.json', 'utf-8', (err, jsonData) => {
//   if (err) {
//      console.log('Error reading File: ', err)
//      throw new Error(err)
//   };
//   const data = JSON.parse(jsonData)
//   const writePath = data.SaveAs
//   //WriteToWorksheet object in json file
//   const sheetObject = data.WriteToWorksheet
//   //return key values in of WriteToWorksheet in json file   
//   for (const key of sheetObject) {
//      //console.log(key)
//      const cellValue = key.WriteToCell; //data values
//      const cellBorder = key.SetCellBorders; //cell or column ranges for borders
//      const bgColor = key.SetCellBackgroundColor; //background fill color used mostly for header range
//      const columnWidth = key.SetColumnWidth; //width formats for columns
//      const cellFont = key.SetCellFont; //font settings
//      const numberFormat = key.SetCellNumberFormat //cell ranges for cells to be set as number format in worksheet
//      const sheetName = key.CreateWorkSheet
//
//      //Initialize Excel workbook
//      XlsxPopulate.fromBlankAsync().then(workbook => {
//         const sheet = workbook.sheet("Sheet1")
//
//         //parse data into workbook
//         for (const data of cellValue) {
//            let cellRef = data.Cell;
//            let cellValue = data.Value;
//            const str = cellRef.replace(/([a-zA-Z0-9-]+):([a-zA-Z0-9-]+)/g, "\"$1\:\$2\"");
//            //console.log('cell values ===>', cellValue)
//            const r = workbook.sheet(0).range(cellRef)
//            r.value([
//               cellValue,
//
//            ]);
//         };
//
//         //setting column widths
//         for (const key of sheetObject) {
//            //console.log(key)
//            //initialize formatting styles
//            const cellBorder = key.SetCellBorders;
//            const columnWidth = key.SetColumnWidth;
//            const cellFont = key.SetCellFont;
//            //set column widths
//            columnWidth.forEach(item => {
//               //const test = item.Cell
//               for (let i = 0; i < columnWidth.length; i++) {
//                  sheet.column(item.Cell).width(item.Width)
//               }
//            });
//         };
//
//         //Formatting fonts
//         cellFont.forEach(item => {
//            //console.log('fonts ====>',item)
//            let cellRange = item.Cell;
//            let fontFamily = item.FontName;
//            let fontSize = item.Size
//            let fontStyle = item.Style
//            const range = sheet.range(item.Cell)
//            range.forEach(item => {
//               item.style({
//                  fontFamily: "Verdana",
//                  fontSize: fontSize,
//                  bold: fontStyle
//               })
//            })
//         });
//
//         //set column borders
//         cellBorder.forEach(item => {
//
//            const range = sheet.range(item.Cell)
//            const borderRange = item.Border
//            console.log('border range ===>', borderRange)
//            range.forEach(item => {
//               if (borderRange == "Top") {
//                  item.style({
//                     topBorder: true
//                  })
//               } else if (borderRange == "Right") {
//                  item.style({
//                     rightBorder: true
//                  })
//               } else if (borderRange == "Left") {
//                  item.style({
//                     leftBorder: true
//                  })
//               } else if (borderRange == "Bottom") {
//                  item.style({
//                     topBorder: true
//                  })
//               }
//            })
//
//
//         });
//
//         //Background fills for cell range
//         bgColor.forEach(item => {
//            const range = sheet.range(item.Cell)
//            range.forEach(item => {
//               item.style("fill", {
//                  type: "solid",
//                  color: {
//                     rgb: "E0E0E0"
//                  }
//
//               })
//            });
//         });
//
//         numberFormat.forEach(item => {
//
//            const range = sheet.range(item.Cell)
//            range.style({
//               numberFormat: true
//            })
//         })
//
//         // Write to file.
//         return workbook.toFileAsync("./out.xlsx");
//
//      })
//   };
//})