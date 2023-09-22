  function createInputDiv(index) {
        const fragment = document.createDocumentFragment();

        const input = document.createElement('input');
        input.classList.add('input-ev');
        input.type = 'text';
        input.id = `ev${index}`;
        input.name = `ev${index}`;
        input.placeholder = index === 0 ? 'Control EV' : `V${index} EV`;

        const div = document.createElement('div');
        div.classList.add('spec', 'input-div');
        div.appendChild(input);

        const span = document.createElement('span');
        span.classList.add('focus-border');

        const i = document.createElement('i');
        span.appendChild(i);

        div.appendChild(span);

        fragment.appendChild(div);

        return fragment;
      }

      const inputsDiv = document.querySelector('.inputs');

      for (let i = 0; i < 17; i++) {
        const inputDiv = createInputDiv(i);

        if (i % 3 === 0) {
          const rowDiv = document.createElement('div');
          rowDiv.classList.add('row');
          inputsDiv.appendChild(rowDiv);
        }

        const rowDiv = inputsDiv.lastElementChild;
        rowDiv.appendChild(inputDiv);
      }

      function readFile(files) {
        try {
          validateInputs();
        } catch (error) {
          alert(error.message);
          location.reload();
        }

        const fileNameDiv = document.querySelector('.file-name'),
          fileDiv = document.querySelector('.file-div'),
          filePromises = [];

        for (const file of files) {
          const fileName = file.name,
            size = file.size,
            fileSize = (size / 1000).toFixed(2),
            fileNameAndSize = `${fileName} - ${addCommas(fileSize)} KB`;

          const reader = new FileReader();

          const filePromise = new Promise((resolve, reject) => {
            reader.onload = (event) => {
              try {
                const fileData = event.target.result,
                  wb = XLSX.read(fileData, { type: 'binary' }),
                  xlRowObj = XLSX.utils.sheet_to_row_object_array(
                    wb.Sheets[wb.SheetNames[0]],
                    { range: 11 }
                  );
                const rawData = xlRowObj.map((a) => ({ ...a }));
                resolve(rawData);
              } catch (error) {
                reject(error);
              }
            };

            reader.readAsBinaryString(file);
          });

          filePromises.push(filePromise);

          if (fileNameDiv.textContent === 'No files selected!') {
            fileNameDiv.textContent = fileNameAndSize;
          } else {
            const p = document.createElement('p');
            p.className = 'file-name';
            p.textContent = fileNameAndSize;
            fileDiv.appendChild(p);
          }
        }

        Promise.all(filePromises)
          .then((rawDataArrays) => {
            const flattenedRawData = rawDataArrays.flat();
            createReport(flattenedRawData);
            downloadXL();
          })
          .catch((error) => {
            alert(error.message);
            location.reload();
          });
      }

      // async function countClicks(endpoint) {
      //   const response = await fetch(
      //     `https://api.countapi.xyz/${endpoint}/gantoniou/gaprepgentrig`
      //   );
      //   const data = await response.json();

      //   if (endpoint === 'info') {
      //     document.getElementById('download-counter').textContent = data.value;
      //     document.getElementById('effort-counter').textContent = roundToHours(
      //       data.value * 8
      //     );
      //   }
      // }

      // countClicks('info');

      function addCommas(num) {
        if (typeof num !== 'number') {
          return num;
        }
        return num.toString().replace(/\B(?<!\.\d*)(?=(\d{3})+(?!\d))/g, ',');
      }

      function roundToHours(mins) {
        return Math.floor(mins / 60);
      }

      function transposeArray(array) {
        const [row] = array;
        return row.map((value, column) => array.map((row) => row[column]));
      }

      const inputs = [];
      let reportHeader = [];

      function validateInputs() {
        const inputsEV = document.querySelectorAll('.input-ev'),
          header = document.querySelector('.input-report').value.trim();

        if (header === '') {
          throw new Error('Please fill in the Report Header!');
        }

        const headerArr = header.split(/\r?\n|\r|\n/g);
        reportHeader.push(headerArr);
        reportHeader = transposeArray(reportHeader);

        for (const input of inputsEV) {
          const trimmedInput = input.value.trim();
          if (trimmedInput.length) {
            const inputArr = trimmedInput.split(/[ ,.]+/).filter(Boolean);
            inputs.push(inputArr);
          }
        }
      }

      const finalResults = [];
      let rawDataCombined = [];

      function createReport(data) {
        const emptyProp = '__EMPTY';
        const deliveredProp = 'Delivered';
        const openedProp = 'Opened';
        const clickedProp = 'Clicked';
        const visitsProp = 'Visits';
        const ordersProp = 'Orders';
        const revenueProp = 'Revenue (Gross Merchandise)';
        const netDemandProp = 'Net Demand [Approved]';

        for (const [i, input] of inputs.entries()) {
          if (!finalResults[i]) {
            finalResults[i] = {
              ' ': '',
              [deliveredProp]: 0,
              [openedProp]: 0,
              [clickedProp]: 0,
              [visitsProp]: 0,
              [ordersProp]: 0,
              [revenueProp]: 0,
              [netDemandProp]: 0,
            };
          }

          for (const d of data) {
            for (let n = 0; n < input.length; n++) {
              if (
                d[emptyProp] &&
                (d[emptyProp].includes(input[n]) ||
                  d[emptyProp].includes(input[n] + '\\'))
              ) {
                rawDataCombined = rawDataCombined.concat(d);
                finalResults[i][' '] = i === 0 ? 'Control' : `V${i}`;
                finalResults[i][deliveredProp] += +d[deliveredProp] || 0;
                finalResults[i][openedProp] += +d[openedProp] || 0;
                finalResults[i][clickedProp] += +d[clickedProp] || 0;
                finalResults[i][visitsProp] += +d[visitsProp] || 0;
                finalResults[i][ordersProp] += +d[ordersProp] || 0;
                finalResults[i][revenueProp] += +d[revenueProp] || 0;
                finalResults[i][netDemandProp] += +d[netDemandProp] || 0;
              }
            }
          }
        }
      }

      function downloadXL(
        fileName = `GapTriggeredFinalReport_${new Date()
          .toJSON()
          .slice(0, 10)
          .replace(/-/g, '')}.xlsx`
      ) {
        if (inputs.length === 0) {
          throw new Error('Missing campaign data!');
        } else {
          const ws1 = XLSX.utils.json_to_sheet(finalResults, { origin: 'A8' }),
            ws2 = XLSX.utils.json_to_sheet(rawDataCombined),
            wb = XLSX.utils.book_new();

          XLSX.utils.book_append_sheet(wb, ws1, 'Results');
          XLSX.utils.book_append_sheet(wb, ws2, 'Raw Data');

          for (let i = 0; i < reportHeader.length; i++) {
            XLSX.utils.sheet_add_aoa(ws1, [reportHeader[i]], {
              origin: `A${i + 1}`,
            });
          }

          XLSX.writeFile(wb, fileName);

         // countClicks('hit');
          alert('Report created successfully!');
          location.reload();
        }
      }
