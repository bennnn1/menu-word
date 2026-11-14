/* menu-word.js
   Reads the locked Word menu template (menu.docx) and turns it into menu data.
   Written in older-style JavaScript on purpose so it also works in the older
   web browsers built into some TVs and signage monitors.

   Any page can use it after loading JSZip:
     <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
     <script src="menu-word.js"></script>

     MenuFromWord.load("menu.docx", function (err, menu) { ... });   // callback style
     var menu = await MenuFromWord.load("menu.docx");                  // or promise style

   menu.csv          -> same columns as the old menu.csv
   menu.community    -> community name typed in the Word file
   menu.week1Starts  -> "YYYY-MM-DD" (the Sunday Week 1 began), or "" if not filled in
   menu.cycleWeeks   -> number of weeks in the Word file
*/
(function () {
  'use strict';
  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var KEYS = ['item', 'appetizer', 'entree', 'sides', 'dessert'];

  function textOf(node) {
    var out = '';
    var kids = node.childNodes;
    for (var i = 0; i < kids.length; i++) {
      var c = kids[i];
      if (c.nodeType !== 1) continue;
      var name = c.localName || c.nodeName.replace(/^.*:/, '');
      if (c.namespaceURI === W && name === 't') out += c.textContent;
      else if (c.namespaceURI === W && name === 'tab') out += ' ';
      else if (c.namespaceURI === W && (name === 'br' || name === 'cr')) out += '\n';
      else if (c.namespaceURI === W && (name === 'delText' || name === 'instrText')) continue;
      else out += textOf(c);
    }
    return out;
  }

  function linesOf(sdt) {
    var pr = sdt.getElementsByTagNameNS(W, 'sdtPr')[0];
    if (pr && pr.getElementsByTagNameNS(W, 'showingPlcHdr').length) return [];
    var content = sdt.getElementsByTagNameNS(W, 'sdtContent')[0];
    if (!content) return [];
    var paras = content.getElementsByTagNameNS(W, 'p');
    var out = [];
    for (var i = 0; i < paras.length; i++) {
      var parts = textOf(paras[i]).split('\n');
      for (var j = 0; j < parts.length; j++) {
        var s = parts[j].replace(/\s+/g, ' ').replace(/^\s+|\s+$/g, '');
        if (s) out.push(s);
      }
    }
    return out;
  }

  function two(n) { return (n < 10 ? '0' : '') + n; }

  // Accepts 10/18/2026, 10/18/26, 2026-10-18 or "October 18, 2026" and backs up to that week's Sunday
  function toSunday(text) {
    var s = String(text || '').replace(/^\s+|\s+$/g, '');
    if (!s) return '';
    var d = null, m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s))) d = new Date(+m[1], m[2] - 1, +m[3]);
    else if ((m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/.exec(s))) d = new Date(m[3].length === 2 ? 2000 + +m[3] : +m[3], m[1] - 1, +m[2]);
    else { var t = Date.parse(s); if (!isNaN(t)) d = new Date(t); }
    if (!d || isNaN(d.getTime())) return '';
    d.setDate(d.getDate() - d.getDay());
    return d.getFullYear() + '-' + two(d.getMonth() + 1) + '-' + two(d.getDate());
  }

  function csvValue(v) {
    v = String(v);
    return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  }

  function parseDocument(xml) {
    var doc = new DOMParser().parseFromString(xml, 'application/xml');
    var data = {}, order = [], community = '', start = '', weeks = 0;
    var sdts = doc.getElementsByTagNameNS(W, 'sdt');
    for (var i = 0; i < sdts.length; i++) {
      var tagEl = sdts[i].getElementsByTagNameNS(W, 'tag')[0];
      var tag = tagEl ? (tagEl.getAttributeNS(W, 'val') || tagEl.getAttribute('w:val')) : '';
      if (!tag) continue;
      var lines = linesOf(sdts[i]);
      if (tag === 'Community') { community = lines.join(' '); continue; }
      if (tag === 'Week1Starts') { start = toSunday(lines.join(' ')); continue; }
      var p = tag.split('|');
      var week = parseInt(String(p[0]).slice(1), 10);
      if (!week || !p[1] || !p[2] || KEYS.indexOf(p[3]) < 0) continue;
      if (week > weeks) weeks = week;
      if (!lines.length) continue;
      var id = week + '|' + p[1] + '|' + p[2];
      if (!data[id]) { data[id] = { week: week, day: p[1], meal: p[2] }; order.push(id); }
      data[id][p[3]] = lines.join('\n');
    }
    var rows = ['Week,Day,Meal,Item,Appetizer,Entree,Sides,Dessert,HolidayGif'];
    for (var k = 0; k < order.length; k++) {
      var r = data[order[k]];
      var cells = [r.week, r.day, r.meal];
      for (var j = 0; j < KEYS.length; j++) cells.push(r[KEYS[j]] || '');
      cells.push('');
      for (var c = 0; c < cells.length; c++) cells[c] = csvValue(cells[c]);
      rows.push(cells.join(','));
    }
    return { csv: rows.join('\n'), community: community, week1Starts: start, cycleWeeks: weeks };
  }

  function load(url, callback) {
    function run(done) {
      var xhr = new XMLHttpRequest();
      xhr.open('GET', url, true);
      xhr.responseType = 'arraybuffer';
      xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200 && !(xhr.status === 0 && xhr.response)) {
          done(new Error('Could not load the Word menu (status ' + xhr.status + ')'));
          return;
        }
        JSZip.loadAsync(xhr.response)
          .then(function (zip) {
            var f = zip.file('word/document.xml');
            if (!f) throw new Error('The menu file is not a Word document');
            return f.async('string');
          })
          .then(function (xml) { done(null, parseDocument(xml)); },
                function (e) { done(new Error('Could not read the Word menu: ' + (e && e.message ? e.message : e))); });
      };
      xhr.send();
    }
    if (callback) { run(callback); return undefined; }
    return new Promise(function (resolve, reject) {
      run(function (err, menu) { if (err) reject(err); else resolve(menu); });
    });
  }

  window.MenuFromWord = { load: load };
})();
