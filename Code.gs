/******************************************************************************/
//Update the sprites through onEdit
//Do not modify
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  switch (sheet.getName()) {
    case "sprites":
      update_screen();
      break;
    case "tiles":
      update_sprites();
      break;
    case "new":
      update_new(sheet);
      break;
  }
}

var color_ss = SpreadsheetApp.getActive().getSheetByName("color palette");
var color_bgs = color_ss.getRange("A91:P106").getBackgrounds();
var color_vals = SpreadsheetApp.getActive().getSheetByName("color palette").getRange("A69:P84").getValues();

function update_colors_var() {
  color_ss = SpreadsheetApp.getActive().getSheetByName("color palette");
  color_bgs = color_ss.getRange("A91:P106").getBackgrounds();
  color_vals = SpreadsheetApp.getActive().getSheetByName("color palette").getRange("A69:P84").getValues();
}

/***************************************************************************** */
//this is for drawing arbitrary sized (but rectangular) sprites
function make_new_sprite() {
  var ss = SpreadsheetApp.getActive().getSheetByName("new");
  var width = parseInt(ss.getRange("A1").getValue());
  var height = parseInt(ss.getRange("A2").getValue());

  if (!(Number.isInteger(width) && Number.isInteger(height))) {
    console.log(`Invalid dimension\nW: ${width}\nH: ${height}`);
    return;
  }

  var origin_row = 5, origin_col = 1;
  var final_row = origin_row + height - 1; var final_col = origin_col + width - 1;

  var arr_name = 'mario_row16'.trim();

  var t_x = parseInt(ss.getRange("C3").getValue());
  var t_y = parseInt(ss.getRange("F3").getValue());
  var s_x = 0; var s_y = 0;
  if(Number.isInteger(t_x))
    s_x = t_x;
  if(Number.isInteger(t_y))
    s_y = t_y;

  var out = `uint16_t ${arr_name}_w = ${width};\n`
    + `uint16_t ${arr_name}_h = ${height};\n`
    + `uint16_t ${arr_name}_x = ${s_x.toString()}, ${arr_name}_y = ${s_y.toString()};\n`
    + `uint8_t ${arr_name}_data[] = {`;
  var pixels = ss.getRange(5, 1, height, width).getValues();
  for (let i = 0; i <= final_row - origin_row; ++i) {
    for (let j = final_col - origin_col; j >= 0; --j) {
      let cell = pixels[i][j];
      if (Number.isInteger(cell)) {
        let c = Math.floor(cell / 16);
        let r = cell % 16;
        out += color_vals[r][c].split(',').reverse().join(',');
      } else {
        out += '0,0,0';
      }
      if (!(i == final_row - origin_row && j == 0)) {
        out += ',';
      }
    }
  }
  out += '};\n'
  out += `sprite ${arr_name}_sprite{${arr_name}_data,${arr_name}_x,`
    + `${arr_name}_y, ${arr_name}_w, ${arr_name}_h, false};\n`
  Logger.log(out);

}

function print_mario(){
  let out = '';
  for(let i = 1; i <=16; i++){
    out += `<< mario_row${i}_sprite`;
  }
  console.log(out);
}

function clearSheet(sheet) {
  var numRows = sheet.getLastRow();
  var numCols = sheet.getLastColumn();
  var range = sheet.getRange(5, 1, numRows, numCols);
  range.setBorder(false, false, false, false, false, false);
  range.setBackground(null);
  //to do: clear everything outside of the current work area
  //range.clearContent();
}


function update_new(ss) {
  var w_cell = ss.getRange("A1");
  var h_cell = ss.getRange("A2");

  var width = parseInt(w_cell.getValue());
  var height = parseInt(h_cell.getValue());

  if (!(Number.isInteger(width) && Number.isInteger(height))) {
    return;
  }

  var range = ss.getRange(5, 1, height, width);
  var static_w = parseInt(ss.getRange("I1").getValue());
  var static_h = parseInt(ss.getRange("I2").getValue());
  var pixels = range.getValues();
  if (width != static_w || height != static_h) {
    clearSheet(ss);
    range.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    ss.getRange("I1").setValue(width);
    ss.getRange("I2").setValue(height);
  }

  var origin_row = 5, origin_col = 1;
  var final_row = origin_row + height - 1; var final_col = origin_col + width - 1;
  for (let i = 0; i <= final_row - origin_row; ++i) {
    for (let j = 0; j <= final_col - origin_col; ++j) {
      let cell = pixels[i][j];
      if (Number.isInteger(cell)) {
        let c = Math.floor(cell / 16);
        let r = cell % 16;
        ss.getRange(i + 5, j + 1).setBackground(color_bgs[r][c]);
      } else {
        ss.getRange(i + 5, j + 1).setBackground(null);
      }
    }
  }
}


/************************************************************************* */
//The vector funcitons are a work in progress and might not work yet
//It's an algorithm to draw a vector that does not have to rely on sending row or column commands
function clear_vect() {
  var ss = SpreadsheetApp.getActive().getSheetByName("sprites");
  ss.getRange("B5:BH82").clearContent();
  ss.getRange(3, 30, 80).setValue(144);
  ss.getRange(39, 1, 1, 60).setValue(144);
  ss.getRange("AD39").setValue(71);
  update_screen();
}

//vectors originate from the origin
function draw_vector() {
  var y0 = 39; //the row of the center point
  var x0 = 30; //the column of the center point
  var col0 = x0 - 10; //change where the end of the vector is
  var row0 = y0 + 22;
  var rowc = row0;
  var colc = col0;
  var a = row0 - y0;
  var b = col0 - x0;
  var c = x0 * row0 - y0 * col0;

  clear_vect();
  var ss = SpreadsheetApp.getActive().getSheetByName("sprites");

  var vect_out = [];
  for (i = 0; i <= row0 - y0; ++i) {
    let temp = [];
    for (j = 0; j <= col0 - x0; ++j)
      temp.push(14);
    vect_out.push(temp);
  }
  while (rowc > y0 || colc > x0) {
    if (rowc == y0)
      colc -= 1;
    else if (colc == x0)
      rowc -= 1;
    else if (a * colc > b * rowc + c) {
      colc -= 1;
    }
    else if (a * colc < b * rowc + c) {
      rowc -= 1;
    }
    else {
      vect_out[rowc - y0 - 1][colc - x0] = 111;
      vect_out[rowc - y0][colc - x0 - 1] = 111;
      colc -= 1;
      rowc -= 1;
    }
    vect_out[rowc - y0][colc - x0] = 111;
  }
  vect_out[row0 - y0][col0 - x0] = 48;
  vect_out[0][0] = 71;
  ss.getRange(3, 30, 80).setValue(144);
  ss.getRange(39, 1, 1, 60).setValue(144);
  ss.getRange("AD39").setValue(71);
  ss.getRange(y0, x0, row0 - y0 + 1, col0 - x0 + 1).setValues(vect_out);
  update_screen();
}


/***************************************************************************************/
//Update screen is called if the main screen is changed
//update sprites is called if one of the smaller sprites are changed
function update_screen() {
  var screen_ss = SpreadsheetApp.getActive().getSheetByName("sprites");
  var color_ss = SpreadsheetApp.getActive().getSheetByName("color palette");
  var colors = color_ss.getRange("A91:P106").getBackgrounds();
  var screen_data = screen_ss.getRange("A3:BH82").getValues();
  var out_colors = [];
  var temp = [];
  var cell = 0;
  for (let i = 0; i < 80; ++i) {
    for (let j = 0; j < 60; ++j) {
      cell = screen_data[i][j];
      if (cell.length < 1) {
        temp.push("#fcfcfc");
        continue;
      }
      temp.push(colors[cell % 16][Math.floor(cell / 16)]);
    }
    out_colors.push(temp);
    temp = [];
  }
  screen_ss.getRange("A3:BH82").setBackgrounds(out_colors);
  screen_ss.getRange("A3:BH82").setBorder(true, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
}

function update_sprites() {
  var tiles_ss = SpreadsheetApp.getActive().getSheetByName("tiles");
  var color_ss = SpreadsheetApp.getActive().getSheetByName("color palette");
  var colors = color_ss.getRange("A91:P106").getBackgrounds();
  var tile_data_8x8 = tiles_ss.getRange("A20:H27").getValues();
  var tile_data_16x16 = tiles_ss.getRange("A2:P17").getValues();
  var tile_data_8x16 = tiles_ss.getRange("A32:H47").getValues();
  var tile_data_8x13 = tiles_ss.getRange("A50:H62").getValues();
  var tile_data_debug = tiles_ss.getRange("A65:G70").getValues();
  var out_colors_8x8 = [];
  var out_colors_16x16 = [];
  var out_colors_8x16 = [];
  var out_colors_8x13 = [];
  var out_colors_debug = [];
  var temp = [];
  var cell = 0;
  //16x16 tile
  for (let i = 0; i < 16; ++i) {
    for (let j = 0; j < 16; ++j) {
      cell = tile_data_16x16[i][j];
      if (cell.length < 1) {
        temp.push("#fcfcfc");
        continue;
      }
      temp.push(colors[cell % 16][Math.floor(cell / 16)]);
    }
    out_colors_16x16.push(temp);
    temp = [];
  }
  //8x8 tile
  temp = [];
  for (let i = 0; i < 8; ++i) {
    for (let j = 0; j < 8; ++j) {
      cell = tile_data_8x8[i][j];
      if (cell.length == 0) {
        temp.push("#fcfcfc");
        continue;
      }
      temp.push(colors[cell % 16][Math.floor(cell / 16)]);
    }
    out_colors_8x8.push(temp);
    temp = [];
  }
  //8x16 tile
  temp = [];
  for (let i = 0; i < 16; ++i) {
    for (let j = 0; j < 8; ++j) {
      cell = tile_data_8x16[i][j];
      if (cell.length == 0) {
        temp.push("#fcfcfc");
        continue;
      }
      temp.push(colors[cell % 16][Math.floor(cell / 16)]);
    }
    out_colors_8x16.push(temp);
    temp = [];
  }
  //8x12 tile
  temp = [];
  for (let i = 0; i < 12; ++i) {
    for (let j = 0; j < 8; ++j) {
      cell = tile_data_8x13[i][j];
      if (cell.length == 0) {
        temp.push("#fcfcfc");
        continue;
      }
      temp.push(colors[cell % 16][Math.floor(cell / 16)]);
    }
    out_colors_8x13.push(temp);
    temp = [];
  }
  //debug tile
  temp = [];
  for (let i = 0; i < 6; ++i) {
    for (let j = 0; j < 7; ++j) {
      cell = tile_data_debug[i][j];
      if (cell.length == 0) {
        temp.push("#fcfcfc");
        continue;
      }
      temp.push(colors[cell % 16][Math.floor(cell / 16)]);
    }
    out_colors_debug.push(temp);
    temp = [];
  }

  tiles_ss.getRange("A20:H27").setBackgrounds(out_colors_8x8);
  tiles_ss.getRange("A2:P17").setBackgrounds(out_colors_16x16);
  tiles_ss.getRange("A32:H47").setBackgrounds(out_colors_8x16);
  tiles_ss.getRange("A50:H61").setBackgrounds(out_colors_8x13);
  tiles_ss.getRange("A65:G70").setBackgrounds(out_colors_debug);

  tiles_ss.getRange("A20:H27").setBorder(true, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  tiles_ss.getRange("A2:P17").setBorder(true, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  tiles_ss.getRange("A32:H47").setBorder(true, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  tiles_ss.getRange("A50:H61").setBorder(true, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  tiles_ss.getRange("A65:G70").setBorder(true, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
}

/*********************************************************************************************/
//Making array from created sprites, and sprite related functions

//Fills the 16x16 and 8x8 sprites with 0s in spaces that are left blank. This is for convenience.
function fill_16x16_tile() {
  var data = SpreadsheetApp.getActive().getSheetByName("tiles").getRange("A2:P17").getValues();
  for (let i = 0; i < 16; ++i)
    for (let j = 0; j < 16; ++j)
      if (data[i][j] == "")
        data[i][j] = 0;
  SpreadsheetApp.getActive().getSheetByName("tiles").getRange("A2:P17").setValues(data);
  update_sprites();
}
function fill_8x8_tile() {
  var data = SpreadsheetApp.getActive().getSheetByName("tiles").getRange("A20:H27").getValues();
  for (let i = 0; i < 8; ++i)
    for (let j = 0; j < 8; ++j)
      if (data[i][j] == "")
        data[i][j] = 0;
  SpreadsheetApp.getActive().getSheetByName("tiles").getRange("A20:H27").setValues(data);
  update_sprites();
}

//generates the COMPRESSED array for 7x6 tiles, rightmost column should be left black (space between chars)
//Compressed tiled do not work (very well) with DMA
function make_7x6_compressed_array() {
  var data = SpreadsheetApp.getActive().getSheetByName("tiles").getRange("A65:G70").getValues();
  var out = "{ ";
  for (let i = 0; i < 6; ++i) {
    for (let j = 6; j >= 0; --j) {
      out += data[i][j];
      if (!(i == 5 && j == 0))
        out += ",";
      else
        out += "};";
    }
  }
  Logger.log(out);
}

//Generates the uncompressed array for the 16x16 sprite
function make_16x16_uncompressed_array() {
  var data = SpreadsheetApp.getActive().getSheetByName("tiles").getRange("A2:P17").getValues();
  var colors = SpreadsheetApp.getActive().getSheetByName("color palette").getRange("A69:P84").getValues();
  var out = "{ ";
  for (let i = 0; i < 16; ++i) {
    for (let j = 15; j >= 0; --j) {
      out += colors[data[i][j] % 16][Math.floor(data[i][j] / 16)];
      if (!(i == 15 && j == 0))
        out += ",";
      else
        out += "};";
    }
  }
  Logger.log(out);
}

//Generates the uncompressed array for the 8x8 sprite
function make_8x8_uncompressed_array() {
  var data = SpreadsheetApp.getActive().getSheetByName("tiles").getRange("A20:H27").getValues();
  var colors = SpreadsheetApp.getActive().getSheetByName("color palette").getRange("A69:P84").getValues();
  var out = "{ ";
  for (let i = 0; i < 8; ++i) {
    for (let j = 7; j >= 0; --j) {
      out += colors[data[i][j] % 16][Math.floor(data[i][j] / 16)];
      if (!(i == 7 && j == 0))
        out += ",";
      else
        out += "};";
    }
  }
  Logger.log(out);
}

//Generates the uncompressed array for the 8x16 sprite
function make_8x16_uncompressed_array() {
  var data = SpreadsheetApp.getActive().getSheetByName("tiles").getRange("A32:H47").getValues();
  var colors = SpreadsheetApp.getActive().getSheetByName("color palette").getRange("A69:P84").getValues();
  var out = "{ ";
  for (let i = 0; i < 16; ++i) {
    for (let j = 7; j >= 0; --j) {
      out += colors[data[i][j] % 16][Math.floor(data[i][j] / 16)].split(',').reverse().join(',');
      if (!(i == 15 && j == 0))
        out += ",";
      else
        out += "};";
    }
  }
  Logger.log(out);
}

//Generates the uncompressed array for the 8x12 sprite
function make_8x12_uncompressed_array() {
  var data = SpreadsheetApp.getActive().getSheetByName("tiles").getRange("A50:H61").getValues();
  var colors = SpreadsheetApp.getActive().getSheetByName("color palette").getRange("A69:P84").getValues();
  var out = "{ ";
  for (let i = 0; i < 12; ++i) {
    for (let j = 7; j >= 0; --j) {
      out += colors[data[i][j] % 16][Math.floor(data[i][j] / 16)];
      if (!(i == 11 && j == 0))
        out += ",";
      else
        out += "};";
    }
  }
  Logger.log(out);
}

/***************************************************************************************/
//Don't modify these

//Generating the main color palette (the big section that starts at A1)
//I ran it and it didn't do anything, so I might have broken it, but it did its job
function color_palette() {
  var primary = 0;
  var bsGR = 0;
  var BgsR = 0;
  var BGrs = 0;
  var light_blue = 0;
  var light_green = 0;
  var light_red = 0;

  var ss = SpreadsheetApp.getActive().getSheetByName("color palette");
  var b = ss.getRange("A1:A64").getValues();
  var g = b;
  var r = b;
  ss.getRange("R" + 65).setBackgroundRGB(132, 64, 16);
  ss.getRange("R" + 65).setValue("16,64,132");
  if (primary) {
    for (let i = 0; i < 64; ++i) { //255 colors
      ss.getRange("B" + (i + 2)).setBackgroundRGB(0, 0, b[i]);
      ss.getRange("B" + (i + 2)).setValue("" + b[i] + ',' + 0 + ',' + 0);
      ss.getRange("C" + (i + 2)).setBackgroundRGB(0, g[i], 0);
      ss.getRange("C" + (i + 2)).setValue("" + 0 + ',' + g[i] + ',' + 0);
      ss.getRange("D" + (i + 2)).setBackgroundRGB(r[i], 0, 0);
      ss.getRange("D" + (i + 2)).setValue("" + 0 + ',' + 0 + ',' + r[i]);
      ss.getRange("E" + (i + 2)).setBackgroundRGB(r[i], g[i], b[i]);
      ss.getRange("E" + (i + 2)).setValue("" + b[i] + ',' + g[i] + ',' + r[i]);
      ss.getRange("F" + (i + 2)).setBackgroundRGB(0, g[i], b[i]);
      ss.getRange("F" + (i + 2)).setValue("" + b[i] + ',' + g[i] + ',' + 0);
      ss.getRange("G" + (i + 2)).setBackgroundRGB(0, g[i], b[63 - i]);
      ss.getRange("G" + (i + 2)).setValue("" + b[63 - i] + ',' + g[i] + ',' + 0);
      ss.getRange("H" + (i + 2)).setBackgroundRGB(r[i], 0, b[i]);
      ss.getRange("H" + (i + 2)).setValue("" + b[i] + ',' + 0 + ',' + r[i]);
      ss.getRange("I" + (i + 2)).setBackgroundRGB(b[i], 0, b[63 - i]);
      ss.getRange("I" + (i + 2)).setValue("" + b[63 - i] + ',' + 0 + ',' + r[i]);
      ss.getRange("J" + (i + 2)).setBackgroundRGB(r[i], g[i], 0);
      ss.getRange("J" + (i + 2)).setValue("" + 0 + ',' + g[i] + ',' + r[i]);
      ss.getRange("K" + (i + 2)).setBackgroundRGB(r[63 - i], g[i], 0);
      ss.getRange("K" + (i + 2)).setValue("" + 0 + ',' + g[i] + ',' + r[63 - i]);
    }
  }
  if (light_blue) {
    for (let i = 0; i < 64; ++i) {
      ss.getRange("O" + (i + 2)).setBackgroundRGB(i * 4, i * 4, 252);
      ss.getRange("O" + (i + 2)).setValue("" + 252 + ',' + i * 4 + ',' + i * 4);
    }
  }
  if (light_green) {
    for (let i = 0; i < 64; ++i) {
      ss.getRange("P" + (i + 2)).setBackgroundRGB(i * 4, 252, i * 4);
      ss.getRange("P" + (i + 2)).setValue("" + i * 4 + ',' + 252 + ',' + i * 4);
    }
  }
  if (light_red) {
    for (let i = 0; i < 64; ++i) {
      ss.getRange("Q" + (i + 2)).setBackgroundRGB(252, i * 4, i * 4);
      ss.getRange("Q" + (i + 2)).setValue("" + i * 4 + ',' + i * 4 + ',' + 252);
    }
  }
  if (bsGR) {
    for (let i = 0; i < 64; ++i) {
      ss.getRange("L" + (i + 2)).setBackgroundRGB(252, 252, i * 4);
      ss.getRange("L" + (i + 2)).setValue("" + 252 + "," + 252 + "," + i * 4);
    }
  }
  if (BgsR) {
    for (let i = 0; i < 64; ++i) {
      ss.getRange("M" + (i + 2)).setBackgroundRGB(252, i * 4, 252);
      ss.getRange("M" + (i + 2)).setValue("" + 252 + "," + i * 4 + "," + 252);
    }
  }
  if (BGrs) {
    for (let i = 0; i < 64; ++i) {
      ss.getRange("N" + (i + 2)).setBackgroundRGB(i * 4, 252, 252);
      ss.getRange("N" + (i + 2)).setValue("" + i * 4 + "," + 252 + "," + 252);
    }
  }
}

//Makes the color pallete array given the values in A69:P84
function make_color_palette_array() {
  var ss = SpreadsheetApp.getActive().getSheetByName("color palette");
  var colors = ss.getRange("A69:P84").getValues();
  var arr = "const uint8_t color_palette[256][3] = {";
  var output = "";
  for (let i = 0; i < 16; ++i) {
    for (let j = 0; j < 16; ++j) {
      output += "{" + colors[j][i].split() + "}";
      if (i != 15 || j != 15)
        output += ",";
    }
  }
  Logger.log(arr + output + "};");
}
