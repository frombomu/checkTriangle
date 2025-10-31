function myFunction() {
var pointer = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var a = Number(pointer.getRange("E13").getValue());
  var b = Number(pointer.getRange("E14").getValue());
  var c = Number(pointer.getRange("E15").getValue());

  if ((a == b) && (b == c)) {
    pointer.getRange("D18").setValue("");
    pointer.getRange("E18").setValue("");
    pointer.getRange("F18").setValue("");
    pointer.getRange("D16").setValue("เป็นสามเหลี่ยม");
    pointer.getRange("E16").setValue("ด้านเท่า");
    pointer.getRange("D17").setValue("สามเหลี่ยมด้านเท่ามีพื้นที่");
    let x = (a * a) * (Math.sqrt(3) / 4);
    pointer.getRange("E17").setValue(x);
    pointer.getRange("F17").setValue("ตารางเซนติเมตร");
    pointer.getRange("D19").setValue("");
    pointer.getRange("E19").setValue("");
    pointer.getRange("F19").setValue("");
  }
  else if (a == b) {
    pointer.getRange("D16").setValue("เป็นสามเหลี่ยม");
    pointer.getRange("E16").setValue("หน้าจั่ว");
    pointer.getRange("D18").setValue("height =");
    let h = pointer.getRange("E18").getValue();
    pointer.getRange("D19").setValue("สามเหลี่ยมหน้าจั่วมีพื้นที่");
    let y = 0.5 * h * c;
    pointer.getRange("E19").setValue(y);
    pointer.getRange("F18").setValue("เซนติเมตร");
    pointer.getRange("F19").setValue("ตารางเซนติเมตร");
    pointer.getRange("D17").setValue("");
    pointer.getRange("E17").setValue("");
    pointer.getRange("F17").setValue("");
  }
  else if (a == c) {
    pointer.getRange("D16").setValue("เป็นสามเหลี่ยม");
    pointer.getRange("E16").setValue("หน้าจั่ว");
    pointer.getRange("D18").setValue("height =");
    let l = pointer.getRange("E18").getValue();
    pointer.getRange("D19").setValue("สามเหลี่ยมหน้าจั่วมีพื้นที่");
    let z = 0.5 * l * b;
    pointer.getRange("E19").setValue(z);
    pointer.getRange("F18").setValue("เซนติเมตร");
    pointer.getRange("F19").setValue("ตารางเซนติเมตร");
    pointer.getRange("D17").setValue("");
    pointer.getRange("E17").setValue("");
    pointer.getRange("F17").setValue("");
  }
  else if (b == c) {
    pointer.getRange("D16").setValue("เป็นสามเหลี่ยม");
    pointer.getRange("E16").setValue("หน้าจั่ว");
    pointer.getRange("D18").setValue("height =");
    let m = pointer.getRange("E18").getValue();
    pointer.getRange("D19").setValue("สามเหลี่ยมหน้าจั่วมีพื้นที่");
    let v = 0.5 * m * a;
    pointer.getRange("E19").setValue(v);
    pointer.getRange("F18").setValue("เซนติเมตร");
    pointer.getRange("F19").setValue("ตารางเซนติเมตร");
    pointer.getRange("D17").setValue("");
    pointer.getRange("E17").setValue("");
    pointer.getRange("F17").setValue("");
  }
  else {
    pointer.getRange("D16").setValue("เป็นสามเหลี่ยม");
    pointer.getRange("E16").setValue("ด้านไม่เท่า");
    pointer.getRange("D18").setValue("angle =");
    let n = pointer.getRange("E18").getValue();
    pointer.getRange("F18").setValue("องศา");

    if (n == 90) {
      pointer.getRange("E19").setValue("เป็นสามเหลี่ยมมุมฉาก");
      pointer.getRange("D20").setValue("สามเหลี่ยมมุมฉากมีพื้นที่");
      let s = (a + b + c) / 2;
      let r = Math.sqrt(s * (s - a) * (s - b) * (s - c));
      pointer.getRange("E20").setValue(r);
      pointer.getRange("F20").setValue("ตารางเซนติเมตร");
      pointer.getRange("D17").setValue("");
      pointer.getRange("E17").setValue("");
      pointer.getRange("F17").setValue("");
      pointer.getRange("D19").setValue("");
      pointer.getRange("F19").setValue("");
    }
    else if ((n < 90) && (n > 0)) {
      pointer.getRange("E19").setValue("เป็นสามเหลี่ยมมุมแหลม");
      pointer.getRange("D20").setValue("สามเหลี่ยมมุมแหลมมีพื้นที่");
      let w = (a + b + c) / 2;
      let q = Math.sqrt(w * (w - a) * (w - b) * (w - c));
      pointer.getRange("E20").setValue(q);
      pointer.getRange("F20").setValue("ตารางเซนติเมตร");
      pointer.getRange("D17").setValue("");
      pointer.getRange("E17").setValue("");
      pointer.getRange("F17").setValue("");
      pointer.getRange("D19").setValue("");
      pointer.getRange("F19").setValue("");
    }
    else if ((n > 90) && (n <= 180)) {
      pointer.getRange("E19").setValue("เป็นสามเหลี่ยมมุมป้าน");
      pointer.getRange("D20").setValue("สามเหลี่ยมมุมป้านมีพื้นที่");
      let t = (a + b + c) / 2;
      let g = Math.sqrt(t * (t - a) * (t - b) * (t - c));
      pointer.getRange("E20").setValue(g);
      pointer.getRange("F20").setValue("ตารางเซนติเมตร");
      pointer.getRange("D17").setValue("");
      pointer.getRange("E17").setValue("");
      pointer.getRange("F17").setValue("");
      pointer.getRange("D19").setValue("");
      pointer.getRange("F19").setValue("");
    }
    else {
      pointer.getRange("E19").setValue("กรอกไม่ถูกต้อง");
      pointer.getRange("D20").setValue("");
      pointer.getRange("E20").setValue("");
      pointer.getRange("F20").setValue("");
    }
  }
}
