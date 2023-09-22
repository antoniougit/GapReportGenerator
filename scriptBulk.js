function createInputDiv() {
        const fragment = document.createDocumentFragment();

        for (let i = 1; i <= 9; i++) {
          const container = document.createElement('div');
          container.classList.add(`input-set`);
          container.id = `input-set${i}`;

          const inputDiv1 = document.createElement('div');
          inputDiv1.classList.add('spec', 'input-div');

          const inputName1 = document.createElement('input');
          inputName1.classList.add('input-text');
          inputName1.type = 'text';
          inputName1.id = `campaign-name${i}`;
          inputName1.name = `campaign-name${i}`;
          inputName1.placeholder = `Campaign Name #${i}`;

          const span1 = document.createElement('span');
          span1.classList.add('focus-border');

          const i1 = document.createElement('i');
          span1.appendChild(i1);

          inputDiv1.appendChild(inputName1);
          inputDiv1.appendChild(span1);

          const inputDiv2 = document.createElement('div');
          inputDiv2.classList.add('spec', 'input-div');

          const inputLink = document.createElement('input');
          inputLink.classList.add('input-text');
          inputLink.type = 'text';
          inputLink.id = `portal-link${i}`;
          inputLink.name = `portal-link${i}`;
          inputLink.placeholder = `Portal Link #${i}`;

          const span2 = document.createElement('span');
          span2.classList.add('focus-border');

          const i2 = document.createElement('i');
          span2.appendChild(i2);

          inputDiv2.appendChild(inputLink);
          inputDiv2.appendChild(span2);

          const inputDiv3 = document.createElement('div');
          inputDiv3.classList.add('spec', 'input-div');

          const inputIds = document.createElement('input');
          inputIds.classList.add('input-text');
          inputIds.type = 'text';
          inputIds.id = `geno-ids${i}`;
          inputIds.name = `geno-ids${i}`;
          inputIds.placeholder = `Geno IDs #${i}`;

          const span3 = document.createElement('span');
          span3.classList.add('focus-border');

          const i3 = document.createElement('i');
          span3.appendChild(i3);

          inputDiv3.appendChild(inputIds);
          inputDiv3.appendChild(span3);

          container.appendChild(inputDiv1);
          container.appendChild(inputDiv2);
          container.appendChild(inputDiv3);

          fragment.appendChild(container);
        }

        document.querySelector('.inputs').appendChild(fragment);
      }

      createInputDiv();

      const file = document.getElementById('file'),
        finalResults = [],
        inputs = [];
      let rawData = [];

      file.addEventListener('change', (e) => {
        const [file] = e.target.files,
          { name: fileName, size } = file,
          fileSize = (size / 1000).toFixed(2),
          fileNameAndSize = `${fileName} - ${addCommas(fileSize)} KB`;
        document.querySelector('.file-name').textContent = fileNameAndSize;
      });

      function readFile(file) {
        try {
          validateInputs();
        } catch (error) {
          alert(error.message);
          location.reload();
        }

        const reader = new FileReader();

        reader.onload = function (e) {
          const fileData = e.target.result,
            wb = XLSX.read(fileData, {
              type: 'binary',
            });
          const xlRowObj = XLSX.utils.sheet_to_row_object_array(
            wb.Sheets[wb.SheetNames[0]]
          );

          rawData = xlRowObj.map((a) => {
            return { ...a };
          });
          createReport(
            rawData.map((a) => {
              return { ...a };
            })
          );
        };
        setTimeout(() => {
          try {
            downloadXL(generateFileName());
          } catch (error) {
            alert(error.message);
            location.reload();
          }
        }, 1000);

        reader.readAsBinaryString(file[0]);
      }

      // function countClicks(endpoint) {
      //   const xhr = new XMLHttpRequest();
      //   xhr.open(
      //     'GET',
      //     `https://api.countapi.xyz/${endpoint}/gantoniou/gaprepgen`
      //   );
      //   xhr.responseType = 'json';
      //   xhr.onload = function () {
      //     if (endpoint === 'info') {
      //       document.getElementById('download-counter').textContent =
      //         this.response.value;
      //       document.getElementById('effort-counter').textContent =
      //         roundToHours(this.response.value * 34);
      //     }
      //   };
      //   xhr.send();
      // }

      // countClicks('info');

      function addCommas(num) {
        return num.toString().replace(/\B(?<!\.\d*)(?=(\d{3})+(?!\d))/g, ',');
      }

      function roundToHours(mins) {
        const hours = Math.floor(mins / 60);
        return hours;
      }

      function validateInputs() {
        const inputSets = [],
          allInputs = document.getElementsByClassName('input-text');
        let inputSetsNum = allInputs.length / 3;

        for (let i = 0; i < inputSetsNum; i++) {
          inputSets.push(
            document
              .getElementById('input-set' + (i + 1))
              .getElementsByTagName('input')
          );
        }

        for (input of allInputs) {
          if (input.value.trim().length) {
            inputs.push(input.value);
          }
        }

        inputSetsNum = inputs.length / 3;

        for (let i = 0; i < inputSetsNum; i++) {
          const [campaignName, portalLink, genoIds] = inputSets[i];
          if (!campaignName.value || !portalLink.value || !genoIds.value) {
            throw new Error('Missing campaign data!');
          } else if (
            !portalLink.value.startsWith('https://portal.persado.com/')
          ) {
            throw new Error('Invalid Portal link!');
          }
        }
      }

      function updateNetDemand(data) {
        data.forEach((d) => {
          d['Net Demand'] =
            d['Revenue (Gross Merchandise)'] -
            (d['Line + Brand Discount (e43) (event43)'] +
              d['Order Level Discount [non-PLCC] (e44) (event44)'] +
              d['PLCC Discount (e48) (event48)'] +
              d['Pay with Points Discount (e329) (event329)']);
        });
      }

      function updateAvgs(results) {
        const avgs = results.reduce((res, current) => {
          return res.concat(Array(3).fill(current));
        }, []);

        for (let i = 0; i < finalResults.length; i++) {
          finalResults[i]['Average Lift Opens'] =
            (
              ((avgs[i][1][1] + avgs[i][2][1]) /
                (avgs[i][1][0] + avgs[i][2][0]) /
                (avgs[i][0][1] / avgs[i][0][0]) -
                1) *
              100
            ).toFixed(2) + '%';
          finalResults[i]['Average Lift Clicks'] =
            (
              ((avgs[i][1][2] + avgs[i][2][2]) /
                (avgs[i][1][0] + avgs[i][2][0]) /
                (avgs[i][0][2] / avgs[i][0][0]) -
                1) *
              100
            ).toFixed(2) + '%';
          finalResults[i]['Average Lift Conversions'] =
            (
              ((avgs[i][1][3] + avgs[i][2][3]) /
                (avgs[i][1][0] + avgs[i][2][0]) /
                (avgs[i][0][3] / avgs[i][0][0]) -
                1) *
              100
            ).toFixed(2) + '%';
          const incRevenue =
            ((avgs[i][1][4] + avgs[i][2][4]) / (avgs[i][1][0] + avgs[i][2][0]) -
              avgs[i][0][4] / avgs[i][0][0]) *
            (avgs[i][1][0] + avgs[i][2][0]);
          finalResults[i]['Incremental Revenue'] =
            incRevenue >= 0
              ? '$' + addCommas(incRevenue.toFixed(2))
              : '-$' + addCommas(Math.abs(incRevenue).toFixed(2));
        }
      }

      function generateFileName() {
        const fileInput = document.getElementById('file'),
          fName = fileInput.files[0].name,
          fileSplit = fName.split('_'),
          fileBrand = fileSplit[0],
          fileDate = fileSplit[fileSplit.length - 1].split('.')[0];
        fileName = `${fileBrand}_${fileDate}.xlsx`;
        return fileName;
      }

      function createReport(rawData) {
        const data = rawData,
          inputObj = [],
          results = [],
          metrics = ['Delivered', 'Opened', 'Clicked', 'Orders', 'Net Demand'],
          variants = ['C', 'T1', 'T2'];

        for (let z = 0; z < inputs.length / 3; z++) {
          inputObj.push({
            'Campaign Name': inputs[z * 3],
            'Portal Link': inputs[z * 3 + 1],
            'Geno Ids': inputs[z * 3 + 2]
              .match(/(?:\d+\.)?\d+/g)
              .filter(Number),
          });
          const input = inputObj[z];
          if (input['Geno Ids'] && input['Geno Ids'].length > 1) {
            let maxIndex = 0;
            for (let i = 1; i < input['Geno Ids'].length; i++) {
              if (input['Geno Ids'][i] > input['Geno Ids'][maxIndex]) {
                maxIndex = i;
              }
            }

            const maxVal = input['Geno Ids'][maxIndex];
            input['Geno Ids'].splice(maxIndex, 1);
            input['Geno Ids'].unshift(maxVal);
          }

          updateNetDemand(data);

          results.push(Array.from({ length: 3 }, () => Array(5).fill(0)));

          const result = results[z];

          for (let i = 0; i < variants.length; i++) {
            const genoId = input['Geno Ids'][i];
            const validEmails = ['PERSXC1X', 'PERSXT1X', 'PERSXT2X'];

            data.forEach((d) => {
              const emailVersion = d['Email version (v11) (evar11)'];
              for (let k = 0; k < metrics.length; k++) {
                if (
                  emailVersion.includes(genoId) &&
                  validEmails.some((email) => emailVersion.includes(email))
                ) {
                  for (let key in d) {
                    if (key.startsWith(metrics[k])) {
                      result[i][k] += d[key];
                      break;
                    }
                  }
                }
                result[i][k] = parseFloat(result[i][k].toFixed(2));
              }
            });

            finalResults.push({
              'Experiment Name & Link': input['Campaign Name'],
              Variant: variants[i],
              Delivered: result[i][0],
              Opened: result[i][1],
              Clicked: result[i][2],
              Orders: result[i][3],
              'Net Demand': result[i][4],
              '': '',
              'Open Rate':
                ((result[i][1] / result[i][0]) * 100).toFixed(2) + '%',
              'Average Lift Opens': 0,
              'Click Rate':
                ((result[i][2] / result[i][0]) * 100).toFixed(2) + '%',
              'Average Lift Clicks': 0,
              'Conversion Rate':
                ((result[i][3] / result[i][0]) * 100).toFixed(2) + '%',
              'Average Lift Conversions': 0,
              'Incremental Revenue': 0,
              'Geno ID': input['Geno Ids'][i],
            });
          }
        }
        updateAvgs(results);
      }

      function downloadXL(fileName) {
        if (inputs.length === 0) {
          throw new Error('Missing campaign data!');
        } else {
          fileName = fileName || 'GapFinalReport.xlsx';
          const ws1 = XLSX.utils.json_to_sheet(finalResults),
            ws2 = XLSX.utils.json_to_sheet(rawData),
            wb = XLSX.utils.book_new();

          XLSX.utils.book_append_sheet(wb, ws1, 'Results');
          XLSX.utils.book_append_sheet(wb, ws2, 'Raw Data');
          const merge = [];
          const fillArray = [9, 9, 9, 9, 9];
          const cArray = [0, 9, 11, 13, 14];

          for (let i = 0; i < fillArray.length; i++) {
            const c = cArray[i];
            for (let j = 0; j < fillArray[i]; j++) {
              merge.push({
                s: { r: 1 + 3 * j, c: c },
                e: { r: 3 + 3 * j, c: c },
              });
            }
          }

          ws1['!merges'] = merge;

          for (let i = 1; i < inputs.length; i += 3) {
            ws1['A' + (i + 1)].l = {
              Target: inputs[i],
            };
          }

          const date = new Date();
          const currentDate = date.toLocaleString('default', {
            month: 'short',
            day: 'numeric',
            year: 'numeric',
          });

          XLSX.utils.sheet_add_aoa(ws1, [['Results Pulled: ' + currentDate]], {
            origin: 'A30',
          });

          XLSX.writeFile(wb, fileName);

          //countClicks('hit');
          alert('Report created successfully!');

          location.reload();
        }
      }
