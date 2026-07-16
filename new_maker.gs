
function new_maker() {
  var ss = SpreadsheetApp.getActive().getSheetByName("new");
  var width = parseInt(ss.getRange("A1").getValue());
  var height = parseInt(ss.getRange("A2").getValue());

  if (!(Number.isInteger(width) && Number.isInteger(height))) {
    console.log(`Invalid dimension\nW: ${width}\nH: ${height}`);
    return;
  }

  //these are used as "global" or "static" variables
  //if the user changes the dimensions, these hold the last values for width/height
  var t_x = parseInt(ss.getRange("C3").getValue());
  var t_y = parseInt(ss.getRange("F3").getValue());
  var s_x = 0; var s_y = 0;
  if (Number.isInteger(t_x))
    s_x = t_x;
  if (Number.isInteger(t_y))
    s_y = t_y;


  //the "raw" pixel data from the mario ss
  var pixels = ss.getRange(5, 1, height, width).getValues();
  var origin_row = 5;
  var origin_col = 1;
  var final_row = origin_row + height - 1;
  var final_col = origin_col + width - 1;

  /*
    sprites is an array in the format:
      [rows, [sprites for that row]]
    where <sprites in that row> is an array in the format:
      [sprite row name/id, sprite data]

    so:
      [rows, [[sprite name/id, sprite data]]]


    to optimize this, I can have the [sprites for that row] array also contain:
    x0, x1, y0, y1 representing the square that contains the sprite
    then i can go and try to merge sprites vertically
    each sprite i can merge saves ~144 bytes from storage space alone (not including the code it takes to add it to the list or whatever)

    it will only save space if two entire sprites can be merged together, otherwise the sprites just get broken up

    or maybe add a way to manually merge sprites
  */
  let sprite_name = ss.getRange("S3").getValue();
  let sprites = [];
  //Logger.log(pixels);
  for (let i = 0; i < pixels.length; ++i) {
    sprites.push(row2sprite(pixels[i], i, sprite_name));
  }

  //the row
  let sprite_names = [];
  let sprite_data = '';
  for (let i = 0; i < sprites.length; ++i) {
    //the sprites in the row
    for (let j = 0; j < sprites[i].length; ++j) {
      //Logger.log(sprites[i][j]);
      sprite_names.push(sprites[i][j][0]);
      sprite_data += sprites[i][j][1] + '\n';
    }
  }

  let sprite_list = `sprite_list ${sprite_name};\n${sprite_name} << `;
  for(let i = 0; i < sprite_names.length; ++i){
    sprite_list += sprite_names[i];
    if(i !== sprite_names.length - 1){
      sprite_list += ' << ';
    }
    else {
      sprite_list += ';';
    }
  }

Logger.log(sprite_data + '\n' + sprite_list);

}


/*
  @param row - list which represents a row
  @param rownum - number of the row, 0-based
  returns a list of strings, representing sprites
*/
function row2sprite(row, rownum, sprite_name) {
  //within the row, x0 is the first item in the list (rightmost)
  let x0 = -1;
  //x1 is the last item (leftmost item)
  let x1 = -1;
  let y0 = rownum;
  let sprites_in_row_count = 0;
  let sprites = [];

  //Logger.log(`ROWNUM: ${rownum}`);
  //need to assemble the sprite right to left
  for (let i = row.length - 1; i >= 0; --i) {
    //convert the list item into an integer, if it exists (just in case)
    if (!!row[i]) {
      row[i] = Number.parseInt(row[i]);
    }

    //this is the first number we've seen
    if (!!row[i] && x0 === -1) {
      x0 = i;
    }

    //there have been 1 or more valid cells and we just encountered an empty cell
    if (!row[i] && x0 !== -1 && x1 === -1) {
      x1 = i + 1;
    }

    //we made it to the end of the row, found a valid start cell but not an end cell
    //that means 0 is the index of the final cell
    if (i === 0 && x0 !== -1 && x1 === -1) {
      x1 = 0;
    }
    //Logger.log(`x0: ${x0}\nx1: ${x1}\n`);

    //found a valid subarray
    if (x0 !== -1 && x1 !== -1) {
      //Logger.log(`x0: ${x0}\nx1: ${x1}\n`);
      //break;
      //for splice, the number of elements in this sub array
      let w = x0 - x1 + 1;
      sprites.push(arr2sprite(row.splice(x1, w), x1, rownum, sprites_in_row_count++, w, sprite_name));
      x0 = -1;
      x1 = -1;
    }

  }

  return sprites;

}

/*
  helper function for row2sprite
  row2sprite will find subarrays within the row and pass each one to arr2sprite
  arr2sprite will be what actually makes the sprite
  width could be recalculated but we have already calculated it before so screw that
  returns a list of strings, in the format [sprite_name, sprite info]
*/
function arr2sprite(arr, x0, y0, sprite_num, width, sprite_name) {
  let sprite = '';

  let this_sprite_name = `${sprite_name}_row${y0}_${sprite_num}`;
  let data_str = `uint8_t ${sprite_name}_row${y0}_${sprite_num}_data = {`;

  let data = '';
  //turn the pixel colors into their BGR values form the color LUT
  for (let i = arr.length - 1; i >= 0; --i) {
    let c = Math.floor(arr[i] / 16);
    let r = arr[i] % 16;
    data_str += color_vals[r][c];
    if (i !== 0) {
      data_str += ", ";
    }
  }

  data_str += "};";

  let sprite_str = `sprite ${this_sprite_name}_sprite{ ${this_sprite_name}_data, ${x0}, ${y0}, ${width}, 1, dma_incr_modes::inc};`;

  let out = `${data_str}\n${sprite_str}\n`

  return [this_sprite_name, out];
}




