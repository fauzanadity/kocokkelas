function processFile() {
  const fileInput = document.getElementById("inputFile").files[0];
  if (!fileInput) {
    alert("Please select an Excel file first");
    return;
  }

  const reader = new FileReader();
  reader.onload = (event) => {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
    createGroups(rows);
  };
  reader.readAsArrayBuffer(fileInput);
}

function createGroups(rows) {
  rows.shift(); // Remove the header row

  // Sort rows alphabetically by name
  rows.sort((a, b) => a[0].localeCompare(b[0]));

  const groups = Array.from({ length: 12 }, () => []);
  const genderCounts = Array.from({ length: 12 }, () => ({ P: 0, L: 0 }));

  rows.forEach(([name, gender], index) => {
    const groupIndex = index % 12;
    groups[groupIndex].push({ name, gender });
    genderCounts[groupIndex][gender]++;
  });

  balanceGroups(groups, genderCounts);

  displayGroups(groups);
}

function balanceGroups(groups, genderCounts) {
  const targetRatio = 1; // Equal P and L
  let adjustments = true;

  while (adjustments) {
    adjustments = false;

    for (let i = 0; i < 12; i++) {
      const total = genderCounts[i].P + genderCounts[i].L;
      const ratio = genderCounts[i].P / (genderCounts[i].L || 1);

      if (ratio > targetRatio + 0.1 || ratio < targetRatio - 0.1) {
        for (let j = 0; j < 12; j++) {
          if (i === j) continue;

          const swapGender = ratio > targetRatio ? "P" : "L";
          const oppositeGender = swapGender === "P" ? "L" : "P";

          const swapIndex = groups[i].findIndex(
            (person) => person.gender === swapGender
          );
          const oppositeIndex = groups[j].findIndex(
            (person) => person.gender === oppositeGender
          );

          if (swapIndex > -1 && oppositeIndex > -1) {
            const temp = groups[i][swapIndex];
            groups[i][swapIndex] = groups[j][oppositeIndex];
            groups[j][oppositeIndex] = temp;

            genderCounts[i][swapGender]--;
            genderCounts[i][oppositeGender]++;
            genderCounts[j][swapGender]++;
            genderCounts[j][oppositeGender]--;

            adjustments = true;
          }
        }
      }
    }
  }
}

function displayGroups(groups) {
  const output = document.getElementById("output");
  output.innerHTML = "";

  groups.forEach((group, index) => {
    const groupDiv = document.createElement("div");
    groupDiv.className = "group";
    groupDiv.innerHTML = `<h2>Group ${index + 1}</h2><ul>${group
      .map((person) => `<li>${person.name} (${person.gender})</li>`)
      .join("")}</ul>`;
    output.appendChild(groupDiv);
  });
}
