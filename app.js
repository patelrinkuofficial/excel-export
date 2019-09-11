let express = require('express');
let app = express();

var xl = require('excel4node');

app.get('/excel',(req, res)=>{
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('Sheet 1');
    // Create a reusable style
    var style = wb.createStyle({
        font: {
            color: '#000000',
            size: 11,
            name: 'Calibri',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var heading = wb.createStyle({
        font: {
            color: '#000000',
            size: 20,
            name: 'Calibri',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
        fill: {
            type: 'pattern', // the only one implemented so far.
            patternType: 'solid', // most common.
            fgColor: 'FFFF00',
        },
        border: {
            left: {
                style: 'thin',
                color: '000000'
            },
            right: {
                style: 'thin',
                color: '000000'
            },
            top: {
                style: 'thin',
                color: '000000'
            },
            bottom: {
                style: 'thin',
                color: '000000'
            }
        }
    });

    var subheading = wb.createStyle({
        font: {
            color: '#000000',
            size: 11,
            name: 'Calibri',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
        fill: {
            type: 'pattern', // the only one implemented so far.
            patternType: 'solid', // most common.
            fgColor: '00B0F0',
        },
        border: {
            left: {
                style: 'thin',
                color: '000000'
            },
            right: {
                style: 'thin',
                color: '000000'
            },
            top: {
                style: 'thin',
                color: '000000'
            },
            bottom: {
                style: 'thin',
                color: '000000'
            }
        }
    });

    // Heading
    ws.cell(1, 1).string('productname').style(heading);
    ws.cell(1, 2).string('producttagline').style(heading);
    ws.cell(1, 3).string('productdescription').style(heading);
    ws.cell(1, 4).string('sku').style(heading);
    ws.cell(1, 5).string('productimage').style(heading);
    ws.cell(1, 6).string('visible').style(heading);
    ws.cell(1, 7).string('branding_template').style(heading);
    ws.cell(1, 8).string('iscustomize').style(heading);
    ws.cell(1, 9).string('flag_name').style(heading);
    ws.cell(1, 10).string('cost_flag').style(heading);
    ws.cell(1, 11).string('information_flag').style(heading);
    ws.cell(1, 12).string('information_contain').style(heading);
    ws.cell(1, 13).string('attribute').style(heading);
    ws.cell(1, 14).string('category').style(heading);
    ws.cell(1, 15).string('gallery').style(heading);
    ws.cell(1, 16).string('setup_price').style(heading);
    ws.cell(1, 17).string('unbranded_price').style(heading);
    ws.cell(1, 18).string('agency_name').style(heading);

    //Sub heading
    ws.cell(2, 1, 3, 1,true).string('Product Name').style(subheading); //productname
    ws.cell(2, 2, 3, 2,true).string('Tag Line').style(subheading); //producttagline
    ws.cell(2, 3, 3, 3,true).string('Description').style(subheading); //productdescription
    ws.cell(2, 4, 3, 4,true).string('SKU').style(subheading); //sku
    ws.cell(2, 5, 3, 5,true).string('').style(subheading); //productimage
    ws.cell(2, 6, 3, 6,true).string('1 for Publish, 0 For private').style(subheading); //visible
    ws.cell(2, 7, 3, 7,true).string('').style(subheading); //branding_template
    ws.cell(2, 8, 3, 8,true).string('Customize for 1 and otherwise  0').style(subheading); //iscustomize
    ws.cell(2, 9, 3, 9,true).string('Unbranded Flag Name').style(subheading); //flag_name
    ws.cell(2, 10, 3, 10,true).string('Addiation Cost Flag name').style(subheading); //cost_flag
    ws.cell(2, 11, 3, 11,true).string('Information  Flag name').style(subheading); //information_flag
    ws.cell(2, 12, 3, 12,true).string('Information  Flag name').style(subheading); //information_contain
    ws.cell(2, 13, 3, 13,true).string('Ex: Color>Blue|Color>Pink').style(subheading); //attribute
    ws.cell(2, 14, 3, 14,true).string('Ex: Pens>Deluxe|Pens>Metal').style(subheading); //category
    ws.cell(2, 15, 3, 15,true).string('').style(subheading); //gallery
    ws.cell(2, 16, 3, 16,true).string('setup_name>unit_price>setup_price|setup_name>unit_price>setup_price').style(subheading); //setup_price
    ws.cell(2, 17, 3, 17,true).string('qty>price>wholeseller_markup>retailer_markup').style(subheading); //unbranded_price
    ws.cell(2, 18, 3, 18,true).string('Agency Name').style(subheading); //agency_name

    // data Set
    ws.cell(5, 1).string('Vertical Flag Set - Small').style(style);
    ws.cell(5, 2).string('Vertical Flag Set - Small').style(style);
    ws.cell(5, 3).string('"Vertical flag – Small (2.4m, Single Sided or double sided, Remember – We can fully customise these flags to suit your office *Must comply with agency brand guidelines<br><br><b>Colors :</b>Sublimated<br><br><b>Measurment :</b>Dimension:1.91 x 0.6m"').style(style);
    ws.cell(5, 4).string('Vertical Flag Set - Small').style(style);
    ws.cell(5, 5).string('VERTICAL-FLAG_1').style(style);
    ws.cell(5, 6).string('1').style(style);
    ws.cell(5, 7).string('').style(style);
    ws.cell(5, 8).string('1').style(style);
    ws.cell(5, 9).string('SINGLE SIDED PRINT').style(style);
    ws.cell(5, 10).string('OPTIONAL EXTRAS').style(style);
    ws.cell(5, 11).string('INFORMATION').style(style);
    ws.cell(5, 12).string(' <p>Prices are in AUD and exclude GST.</p> <p>Freight on all portal orders is to be confirmed upon final quote</p>').style(style);
    ws.cell(5, 13).string('').style(style);
    ws.cell(5, 14).string('Century 21 Estate Generic>Signs & Flags').style(style);
    ws.cell(5, 15).string('VERTICAL-FLAG-1|VERTICAL-FLAG_11').style(style);
    ws.cell(5, 16).string('Double Sided Print>20>0|Spike Base>1>29|Wall Mount>29>0|Cross Base>37|Vehicle Base>40>0|Square Metal Base>40>0|Water Bag (12kg)>12>0').style(style);
    ws.cell(5, 17).string('1>85.00>0>0').style(style);
    ws.cell(5, 18).string('Century 21').style(style);
    
    wb.write('excel/Excel'+'2'+'.xlsx');
    res.json('1')
});

app.listen(3000,()=>{
    console.log('server start');
});