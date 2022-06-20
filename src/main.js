/* eslint-disable no-return-assign */
/* const XLSX = require("xlsx"); */

document.addEventListener('DOMContentLoaded', () => {
  const aimsInput = document.querySelector('.aims-input');
  const dutValue = document.querySelector('.dutflowresult');
  const refValue = document.querySelector('.refflowresult');
  const devValue = document.querySelector('.deviationresult');
  const flowRangeInput = document.querySelector('.flowrange');
  const aimsSwitcher = document.querySelector('.aims-switcher-checkbox');
  const table = document.querySelector('.tableBody');
  let aims = [100, 75, 50, 25, 10];
  let dutFlow = 0;
  let refFlow = 0;
  let flowRange;

  aimsSwitcher.addEventListener('change', () => {
    aimsInput.toggleAttribute('disabled');
    if (!aimsSwitcher.checked) {
      aims = [100, 75, 50, 25, 10];
      calcTable();
    }
  });


  flowRangeInput.addEventListener('input', () => {
    flowRange = +flowRangeInput.value;
    calcTable();
  });

  aimsInput.addEventListener('input', (e) => {
    if (e.target.value.match(/[^0-9,]/)) {
      e.target.value = e.target.value.slice(0, -1);
    } else {
      aims = aimsInput.value.split(',');
    }
    calcTable();
  });

  function calcTable() {
    document.querySelectorAll('.dutflowresult-value').forEach(e => e.remove());
    document.querySelectorAll('.refflowresult-value').forEach(e => e.remove());
    document.querySelectorAll('.deviationresult-value').forEach(e => e.remove());
    document.querySelectorAll('.row').forEach(e => e.remove());
    for (let i = 0; i < aims.length; i++) {
      for (let j = 0; j < 3; j++) {
        const row = document.createElement('tr');
        row.className = 'row';
        row.classList.add(`rowDut${i}`);

        const dutColumn = document.createElement('td');
        dutColumn.className = 'dutColumn';
        const refColumn = document.createElement('td');
        refColumn.className = 'refColumn';
        const devColumn = document.createElement('td');
        devColumn.className = 'devColumn';

        const divDut = document.createElement('div');
        divDut.className = 'dutflowresult-value';
        const divRef = document.createElement('div');
        divRef.className = 'refflowresult-value';
        const divDev = document.createElement('div');
        divDev.className = 'deviationresult-value';
        calcDutFlow();
        calcRefFlow();
        const dutFlowI = +(dutFlow * (aims[i] / 100)).toFixed(4);
        const refFlowI = +(refFlow * (aims[i] / 100)).toFixed(4);
        const deviation = (((dutFlowI - refFlowI) / flowRange) * 100).toFixed(2);
        divDut.innerHTML = `${dutFlowI}`;
        dutColumn.innerHTML = `${dutFlowI}`;
        dutValue.appendChild(divDut);
        divRef.innerHTML = `${refFlowI}`;
        refColumn.innerHTML = `${refFlowI}`;
        refValue.appendChild(divRef);
        divDev.innerHTML = `${deviation}`;
        devColumn.innerHTML = `${deviation}`;
        devValue.appendChild(divDev);
        table.appendChild(row);
        row.appendChild(dutColumn);
        row.appendChild(refColumn);
        row.appendChild(devColumn);
      }
    }
  }

  function calcDutFlow() {
    return dutFlow = +(Math.random() * ((flowRange + (flowRange * 0.0005)) - (flowRange - (flowRange * 0.0005))) + (flowRange - (flowRange * 0.0005)));
  }

  function calcRefFlow() {
    return refFlow = +(Math.random() * ((flowRange + (flowRange * 0.001)) - (flowRange - (flowRange * 0.001))) + (flowRange - (flowRange * 0.001)));
  }

  document.getElementById('sheetjsexport').addEventListener('click', () => {
    const wb = XLSX.utils.table_to_book(document.getElementById('TableToExport'));
    XLSX.writeFile(wb, 'CalibrationSheet.xlsx');
  });

});
