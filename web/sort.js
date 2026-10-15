(function () {
  "use strict";

  var table = document.getElementById("resultTable");
  var body = document.getElementById("resultBody");
  var buttons;
  var sortColumn = -1;
  var sortDirection = 1;

  if (!table || !body) { return; }

  function cellText(row, index) {
    var cell = row.cells && row.cells[index];
    return cell ? String(cell.textContent !== undefined ? cell.textContent : cell.innerText || "").replace(/^\s+|\s+$/g, "") : "";
  }

  function toNumber(value) {
    var number = parseFloat(String(value).replace(/,/g, ""));
    return isNaN(number) ? 0 : number;
  }

  function toDateValue(value) {
    var digits = String(value || "").replace(/[^0-9]/g, "");
    var number = parseFloat(digits);
    return isNaN(number) ? 0 : number;
  }

  function compareValues(left, right, type) {
    if (type === "number") { return toNumber(left) - toNumber(right); }
    if (type === "date") { return toDateValue(left) - toDateValue(right); }
    return String(left).localeCompare(String(right));
  }

  function updateMarks(activeButton) {
    var i, mark;
    for (i = 0; i < buttons.length; i += 1) {
      mark = buttons[i].getElementsByTagName("span")[0];
      if (buttons[i] === activeButton) {
        buttons[i].className = "sort-button active";
        if (mark) { mark.innerText = sortDirection === 1 ? "▲" : "▼"; }
        buttons[i].parentNode.setAttribute("aria-sort", sortDirection === 1 ? "ascending" : "descending");
      } else {
        buttons[i].className = "sort-button";
        if (mark) { mark.innerText = "↕"; }
        buttons[i].parentNode.setAttribute("aria-sort", "none");
      }
    }
  }

  function sortRows(button) {
    var column = parseInt(button.getAttribute("data-sort-column"), 10);
    var type = button.getAttribute("data-sort-type") || "text";
    var rows = [];
    var i, row;

    for (i = 0; i < body.rows.length; i += 1) {
      row = body.rows[i];
      if ((" " + row.className + " ").indexOf(" empty-row ") < 0) { rows.push(row); }
    }
    if (!rows.length) { return; }

    if (sortColumn === column) {
      sortDirection = sortDirection * -1;
    } else {
      sortColumn = column;
      sortDirection = (type === "number" || type === "date") ? -1 : 1;
    }

    rows.sort(function (leftRow, rightRow) {
      var result = compareValues(cellText(leftRow, column), cellText(rightRow, column), type);
      if (result === 0 && column !== 0) {
        result = compareValues(cellText(leftRow, 0), cellText(rightRow, 0), "number");
      }
      return result * sortDirection;
    });

    for (i = 0; i < rows.length; i += 1) { body.appendChild(rows[i]); }
    updateMarks(button);
  }

  buttons = table.getElementsByTagName("button");
  for (var i = 0; i < buttons.length; i += 1) {
    if ((" " + buttons[i].className + " ").indexOf(" sort-button ") >= 0) {
      buttons[i].onclick = (function (button) {
        return function () { sortRows(button); };
      }(buttons[i]));
    }
  }
}());