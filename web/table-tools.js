// Adds a client-side search box and clickable sortable column headers to
// every <table class="table"> on the page. Purely client-side - it only
// reorders/hides rows already rendered by the server, no data is refetched.
(function () {
  function cellText(cell) {
    return cell.textContent.trim().toLowerCase();
  }

  function sortTable(table, colIndex, ascending) {
    var tbody = table.tBodies[0];
    var rows = Array.prototype.slice.call(tbody.rows);
    rows.sort(function (a, b) {
      var av = cellText(a.cells[colIndex]);
      var bv = cellText(b.cells[colIndex]);
      if (av < bv) return ascending ? -1 : 1;
      if (av > bv) return ascending ? 1 : -1;
      return 0;
    });
    rows.forEach(function (row) { tbody.appendChild(row); });
  }

  function addSortHandlers(table) {
    var headerRow = table.tHead && table.tHead.rows[0];
    if (!headerRow) return;

    Array.prototype.forEach.call(headerRow.cells, function (th, index) {
      if (!th.textContent.trim()) return; // skip empty/action columns

      th.style.cursor = 'pointer';
      th.dataset.sortDir = '';

      th.addEventListener('click', function () {
        var ascending = th.dataset.sortDir !== 'asc';

        Array.prototype.forEach.call(headerRow.cells, function (sib) {
          sib.dataset.sortDir = '';
          sib.textContent = sib.textContent.replace(/ [▲▼]$/, '');
        });

        th.dataset.sortDir = ascending ? 'asc' : 'desc';
        th.textContent = th.textContent.replace(/ [▲▼]$/, '') + (ascending ? ' ▲' : ' ▼');

        sortTable(table, index, ascending);
      });
    });
  }

  function addSearchBox(table) {
    var wrapper = document.createElement('div');
    wrapper.className = 'mb-2';

    var input = document.createElement('input');
    input.type = 'search';
    input.className = 'form-control';
    input.placeholder = 'Search...';
    wrapper.appendChild(input);

    table.parentNode.insertBefore(wrapper, table);

    input.addEventListener('input', function () {
      var query = input.value.trim().toLowerCase();
      var rows = table.tBodies[0].rows;
      Array.prototype.forEach.call(rows, function (row) {
        var text = row.textContent.toLowerCase();
        row.style.display = text.indexOf(query) === -1 ? 'none' : '';
      });
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    var tables = document.querySelectorAll('table.table');
    Array.prototype.forEach.call(tables, function (table) {
      if (!table.tHead || !table.tBodies.length) return;
      addSearchBox(table);
      addSortHandlers(table);
    });
  });
})();
