let flights = [];

//============================
// Load Excel
//============================

document.getElementById("excelFile").addEventListener("change", function (e) {

    const file = e.target.files[0];
    if (!file) return;

    document.getElementById("fileName").innerHTML = file.name;

    const reader = new FileReader();

    reader.onload = function (ev) {

        const workbook = XLSX.read(ev.target.result, {
            type: "array"
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: ""
        });

        parseRows(rows);

    };

    reader.readAsArrayBuffer(file);

});

//============================
// Parse
//============================

function parseRows(rows) {

    flights = [];

    let current = null;

    for (let i = 7; i < rows.length; i++) {

        const r = rows[i];

        const flight = r[1];
        const crew = r[16];

        if (flight !== "" && flight != null) {

            current = {

                flight: String(flight).trim(),
                crew: []

            };

            flights.push(current);

        }

        if (!current) continue;

        if (!crew) continue;

        current.crew.push(parseCrew(crew));

    }

    document.getElementById("status").innerHTML =
        "Đã tải <b>" + flights.length + "</b> chuyến.";

    console.log(flights);

}

//============================
// Parse Crew
//============================

function parseCrew(text) {

    text = String(text).trim();

    let role = "";

    const m = text.match(/\((.*?)\)/);

    if (m)
        role = m[1].replace("+", "");

    text = text.replace(/\(.*?\)/, "");

    text = text.replace(/^\-/, "");

    text = text.replace(/\d+$/, "");

    text = text.trim();

    return {

        role: role,
        name: text

    };

}

//============================
// Search
//============================

document.getElementById("searchBtn").onclick = search;

document.getElementById("flightInput").addEventListener("keydown", function (e) {

    if (e.key === "Enter")
        search();

});

function search() {

    let keyword = document.getElementById("flightInput").value;

    keyword = keyword.replace("VJ", "").trim();

    const flight = flights.find(f => f.flight == keyword);

    if (!flight) {

        document.getElementById("result").innerHTML =
            '<div class="notfound">Không tìm thấy chuyến bay.</div>';

        return;

    }

    let html = "";

    html += '<div class="card">';
    html += '<div class="card-header">';
    html += "VJ" + flight.flight;
    html += "</div>";
    html += '<div class="card-body">';

    flight.crew.forEach(c => {

        const list = [];

        flights.forEach(f => {

            if (f.crew.some(x => x.name == c.name))
                list.push("VJ" + f.flight);

        });

        html += `

        <div class="member">

            <div class="member-left">

                <div class="role">${c.role}</div>

                <div class="name">${c.name}</div>

            </div>

            <div class="flights">

                ${list.join(" • ")}

            </div>

        </div>

        `;

    });

    html += "</div></div>";

    document.getElementById("result").innerHTML = html;

}
