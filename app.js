// =========================================================
// CREW LIST - CREW PAIRING CHECK
// =========================================================

let flights = [];


// =========================================================
// ELEMENTS
// =========================================================

const excelFile = document.getElementById("excelFile");
const fileName = document.getElementById("fileName");
const statusBox = document.getElementById("status");

const departureInput =
    document.getElementById("departureFlight");

const arrivalInput =
    document.getElementById("arrivalFlight");

const checkBtn =
    document.getElementById("checkBtn");

const resultBox =
    document.getElementById("result");


// =========================================================
// LOAD EXCEL
// =========================================================

excelFile.addEventListener("change", function (event) {

    const file = event.target.files[0];

    if (!file) return;

    fileName.textContent = file.name;

    statusBox.innerHTML =
        "⏳ Đang đọc Crew List...";

    const reader = new FileReader();

    reader.onload = function (e) {

        try {

            const workbook = XLSX.read(
                e.target.result,
                {
                    type: "array",
                    cellDates: false
                }
            );

            const sheet =
                workbook.Sheets[
                    workbook.SheetNames[0]
                ];

            const rows =
                XLSX.utils.sheet_to_json(
                    sheet,
                    {
                        header: 1,
                        defval: ""
                    }
                );

            parseCrewList(rows);

        }
        catch (error) {

            console.error(error);

            flights = [];

            statusBox.innerHTML =
                "❌ Không thể đọc file Excel.";

        }

    };

    reader.readAsArrayBuffer(file);

});


// =========================================================
// NORMALIZE HEADER
// =========================================================

function normalizeHeader(value) {

    return String(
        value ?? ""
    )
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ")
    .replace(/\n/g, " ");

}


// =========================================================
// FIND COLUMN
// =========================================================

function findColumn(header, names) {

    for (
        let i = 0;
        i < header.length;
        i++
    ) {

        const value =
            normalizeHeader(
                header[i]
            );

        if (!value) continue;

        for (const name of names) {

            if (
                value ===
                normalizeHeader(name)
            ) {

                return i;

            }

        }

    }

    return -1;

}


// =========================================================
// FIND CREW COLUMN
// =========================================================

function findCrewColumn(header) {

    return findColumn(
        header,
        [
            "CREW",
            "CREW NAME",
            "CREW NAMES",
            "CREW LIST",
            "CREW MEMBER",
            "CREW MEMBERS",
            "CREW MEMBERS NAME",
            "TỔ BAY",
            "TEN TO BAY"
        ]
    );

}


// =========================================================
// FIND CREW COUNT COLUMN
// =========================================================

function findCrewCountColumn(header) {

    return findColumn(
        header,
        [
            "CREW #",
            "CREW#",
            "CREW NO",
            "CREW NO.",
            "CREW NUMBER",
            "CREW COUNT",
            "NUMBER OF CREW",
            "CREW MEMBERS #",
            "CREW MEMBER #",
            "SỐ LƯỢNG TỔ BAY",
            "SO LUONG TO BAY"
        ]
    );

}


// =========================================================
// FIND HEADER ROW
// =========================================================

function findHeaderRow(rows) {

    for (
        let i = 0;
        i < Math.min(rows.length, 40);
        i++
    ) {

        const row =
            rows[i] || [];

        let hasFlight = false;
        let hasCrew = false;

        for (const cell of row) {

            const value =
                normalizeHeader(cell);

            if (
                value === "FLT" ||
                value === "FLIGHT" ||
                value === "FLIGHT NO" ||
                value === "FLIGHT NUMBER"
            ) {

                hasFlight = true;

            }

            if (
                value === "CREW" ||
                value === "CREW NAME" ||
                value === "CREW LIST" ||
                value === "CREW MEMBERS" ||
                value === "TỔ BAY"
            ) {

                hasCrew = true;

            }

        }

        if (
            hasFlight &&
            hasCrew
        ) {

            return i;

        }

    }

    return -1;

}


// =========================================================
// PARSE CREW LIST
// =========================================================

function parseCrewList(rows) {

    flights = [];

    const headerRow =
        findHeaderRow(rows);

    if (headerRow === -1) {

        statusBox.innerHTML =
            "❌ Không tìm thấy dòng header FLT / CREW.";

        return;

    }


    const header =
        rows[headerRow];


    // -----------------------------------------------------
    // FIND COLUMNS
    // -----------------------------------------------------

    const flightCol =
        findColumn(
            header,
            [
                "FLT",
                "FLIGHT",
                "FLIGHT NO",
                "FLIGHT NUMBER"
            ]
        );

    const crewCol =
        findCrewColumn(header);

    const crewCountCol =
        findCrewCountColumn(header);

    const dateCol =
        findColumn(
            header,
            [
                "DATE",
                "DAY"
            ]
        );

    const typeCol =
        findColumn(
            header,
            ["TYPE"]
        );

    const regCol =
        findColumn(
            header,
            [
                "REG",
                "REGISTRATION"
            ]
        );

    const acCol =
        findColumn(
            header,
            [
                "AC",
                "A/C",
                "AIRCRAFT"
            ]
        );

    const depCol =
        findColumn(
            header,
            [
                "DEP",
                "FROM"
            ]
        );

    const arrCol =
        findColumn(
            header,
            [
                "ARR",
                "TO"
            ]
        );

    const stdCol =
        findColumn(
            header,
            ["STD"]
        );

    const staCol =
        findColumn(
            header,
            ["STA"]
        );

    const etdCol =
        findColumn(
            header,
            ["ETD"]
        );

    const etaCol =
        findColumn(
            header,
            ["ETA"]
        );


    if (flightCol === -1) {

        statusBox.innerHTML =
            "❌ Không tìm thấy cột FLT.";

        return;

    }


    if (crewCol === -1) {

        statusBox.innerHTML =
            "❌ Không tìm thấy cột CREW.";

        return;

    }


    console.log(
        "HEADER ROW:",
        headerRow + 1
    );

    console.log(
        "FLT COLUMN:",
        columnLetter(flightCol)
    );

    console.log(
        "CREW COLUMN:",
        columnLetter(crewCol)
    );

    console.log(
        "CREW # COLUMN:",
        crewCountCol >= 0
            ? columnLetter(crewCountCol)
            : "Không có"
    );


    // -----------------------------------------------------
    // READ DATA
    // -----------------------------------------------------

    let currentFlight = null;


    for (
        let i = headerRow + 1;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i] || [];


        // =================================================
        // NEW FLIGHT
        // =================================================

        const rawFlight =
            row[flightCol];


        const flightText =
            String(
                rawFlight ?? ""
            ).trim();


        if (flightText !== "") {

            currentFlight = {

                flight:
                    normalizeFlightNumber(
                        flightText
                    ),

                originalFlight:
                    flightText,

                date:
                    dateCol >= 0
                        ? row[dateCol]
                        : "",

                type:
                    typeCol >= 0
                        ? row[typeCol]
                        : "",

                reg:
                    regCol >= 0
                        ? row[regCol]
                        : "",

                aircraft:
                    acCol >= 0
                        ? row[acCol]
                        : "",

                dep:
                    depCol >= 0
                        ? normalizeAirport(
                            row[depCol]
                        )
                        : "",

                arr:
                    arrCol >= 0
                        ? normalizeAirport(
                            row[arrCol]
                        )
                        : "",

                std:
                    stdCol >= 0
                        ? row[stdCol]
                        : "",

                sta:
                    staCol >= 0
                        ? row[staCol]
                        : "",

                etd:
                    etdCol >= 0
                        ? row[etdCol]
                        : "",

                eta:
                    etaCol >= 0
                        ? row[etaCol]
                        : "",

                crewCount:
                    crewCountCol >= 0
                        ? row[crewCountCol]
                        : "",

                crew: [],

                rowIndex: i

            };


            flights.push(
                currentFlight
            );

        }


        if (!currentFlight) {
            continue;
        }


        // =================================================
        // CREW
        // =================================================

        const crewValue =
            row[crewCol];


        if (
            crewValue === undefined ||
            crewValue === null ||
            String(crewValue).trim() === ""
        ) {

            continue;

        }


        const members =
            String(
                crewValue
            )
            .split(/\r?\n/)
            .map(
                x => x.trim()
            )
            .filter(
                x => x !== ""
            );


        members.forEach(
            member => {

                const crew =
                    parseCrew(
                        member
                    );


                if (
                    crew.name
                ) {

                    currentFlight.crew.push(
                        crew
                    );

                }

            }
        );

    }


    // =====================================================
    // REMOVE DUPLICATE CREW
    // =====================================================

    flights.forEach(
        flight => {

            const seen =
                new Set();


            flight.crew =
                flight.crew.filter(
                    crew => {

                        const key =
                            normalizeName(
                                crew.name
                            );


                        if (
                            !key ||
                            seen.has(key)
                        ) {

                            return false;

                        }


                        seen.add(key);

                        return true;

                    }
                );

        }
    );


    // =====================================================
    // STATUS
    // =====================================================

    let totalCrew = 0;

    flights.forEach(
        flight => {

            totalCrew +=
                flight.crew.length;

        }
    );


    statusBox.innerHTML =

        "✅ Đã tải <b>" +
        flights.length +
        "</b> chuyến bay." +

        " &nbsp; " +

        "Tìm thấy <b>" +
        totalCrew +
        "</b> thành viên." +

        "<br>" +

        "<small>" +

        "Header: dòng " +
        (headerRow + 1) +

        " | FLT: cột " +
        columnLetter(flightCol) +

        " | CREW: cột " +
        columnLetter(crewCol) +

        (
            crewCountCol >= 0
                ? " | CREW #: cột " +
                  columnLetter(
                      crewCountCol
                  )
                : ""
        ) +

        "</small>";


    console.log(
        "FLIGHTS:",
        flights
    );

}


// =========================================================
// NORMALIZE FLIGHT NUMBER
// =========================================================

function normalizeFlightNumber(value) {

    let text =
        String(
            value ?? ""
        )
        .trim()
        .toUpperCase();


    text =
        text.replace(
            /^VJ/,
            ""
        );


    text =
        text.replace(
            /\s+/g,
            ""
        );


    if (
        /^\d+\.0$/.test(text)
    ) {

        text =
            text.replace(
                ".0",
                ""
            );

    }


    return text;

}


// =========================================================
// NORMALIZE AIRPORT
// =========================================================

function normalizeAirport(value) {

    return String(
        value ?? ""
    )
    .trim()
    .toUpperCase();

}


// =========================================================
// NORMALIZE NAME
// =========================================================

function normalizeName(value) {

    return String(
        value ?? ""
    )
    .toUpperCase()
    .trim()
    .replace(
        /\s+/g,
        " "
    );

}


// =========================================================
// PARSE CREW
// =========================================================

function parseCrew(text) {

    let value =
        String(
            text ?? ""
        ).trim();


    let role = "";


    // -----------------------------------------------------
    // LẤY ROLE TRONG NGOẶC
    // Ví dụ:
    // NGUYEN VAN A(CP)
    // NGUYEN VAN A(CP, DHD)
    // -----------------------------------------------------

    const matches =
        value.match(
            /\((.*?)\)/g
        );


    if (matches) {

        const roles = [];


        matches.forEach(
            item => {

                let inside =
                    item
                        .replace(
                            /^\(/,
                            ""
                        )
                        .replace(
                            /\)$/,
                            ""
                        )
                        .trim();


                inside =
                    inside
                        .split(",")
                        .map(
                            x => x.trim()
                        )
                        .filter(
                            x =>
                                x &&
                                x.toUpperCase() !==
                                "DHD"
                        )
                        .join(" / ");


                if (inside) {

                    roles.push(
                        inside
                    );

                }

            }
        );


        role =
            roles.join(" / ");

    }


    // -----------------------------------------------------
    // XÓA NGOẶC
    // -----------------------------------------------------

    value =
        value.replace(
            /\(.*?\)/g,
            ""
        );


    // -----------------------------------------------------
    // XÓA +
    // -----------------------------------------------------

    value =
        value.replace(
            /\+/g,
            ""
        );


    value =
        value.replace(
            /^\s*-\s*/,
            ""
        );


    value =
        value.trim();


    return {

        role: role,

        name: value

    };

}


// =========================================================
// FIND FLIGHT
// =========================================================

// =========================================================
// FIND FLIGHT
// =========================================================
// Khi có nhiều chuyến cùng số hiệu:
// Ưu tiên chuyến BAY TỪ HAN
// =========================================================

function findFlight(number, options = {}) {

    const target =
        normalizeFlightNumber(number);

    const candidates =
        flights.filter(
            flight =>
                normalizeFlightNumber(
                    flight.flight
                ) === target
        );


    if (!candidates.length) {

        return null;

    }


    // =====================================================
    // Nếu yêu cầu chuyến đi từ HAN
    // =====================================================

    if (
        options.fromHAN === true
    ) {

        const hanFlight =
            candidates.find(
                flight =>
                    normalizeAirport(
                        flight.dep
                    ) === "HAN"
            );


        if (hanFlight) {

            return hanFlight;

        }

    }


    // =====================================================
    // Nếu chỉ có một chuyến
    // =====================================================

    if (
        candidates.length === 1
    ) {

        return candidates[0];

    }


    // =====================================================
    // Có nhiều chuyến nhưng không xác định được HAN
    // =====================================================

    return candidates[0];

}


// =========================================================
// DATE + TIME
// =========================================================

function getFlightDateTime(
    flight
) {

    if (!flight) {
        return NaN;
    }


    let time =
        flight.etd;


    if (
        time === "" ||
        time === null ||
        time === undefined
    ) {

        time =
            flight.std;

    }


    return parseDateTime(
        flight.date,
        time
    );

}


// =========================================================
// PARSE DATETIME
// =========================================================

function parseDateTime(
    dateValue,
    timeValue
) {

    const date =
        parseDate(
            dateValue
        );


    const time =
        parseTime(
            timeValue
        );


    if (
        !date ||
        !time
    ) {

        return NaN;

    }


    return new Date(

        date.year,
        date.month - 1,
        date.day,

        time.hour,
        time.minute,

        0,
        0

    ).getTime();

}


// =========================================================
// PARSE DATE
// =========================================================

function parseDate(value) {

    if (
        value instanceof Date
    ) {

        return {

            year:
                value.getFullYear(),

            month:
                value.getMonth() + 1,

            day:
                value.getDate()

        };

    }


    if (
        typeof value === "number"
    ) {

        const epoch =
            new Date(
                Date.UTC(
                    1899,
                    11,
                    30
                )
            );


        const date =
            new Date(
                epoch.getTime() +
                value * 86400000
            );


        return {

            year:
                date.getUTCFullYear(),

            month:
                date.getUTCMonth() + 1,

            day:
                date.getUTCDate()

        };

    }


    const text =
        String(
            value ?? ""
        ).trim();


    if (!text) {
        return null;
    }


    let match =
        text.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/
        );


    if (match) {

        let day =
            Number(match[1]);

        let month =
            Number(match[2]);

        let year =
            Number(match[3]);


        if (
            year < 100
        ) {

            year += 2000;

        }


        return {
            year,
            month,
            day
        };

    }


    match =
        text.match(
            /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/
        );


    if (match) {

        return {

            year:
                Number(match[1]),

            month:
                Number(match[2]),

            day:
                Number(match[3])

        };

    }


    const parsed =
        new Date(text);


    if (
        !isNaN(
            parsed.getTime()
        )
    ) {

        return {

            year:
                parsed.getFullYear(),

            month:
                parsed.getMonth() + 1,

            day:
                parsed.getDate()

        };

    }


    return null;

}


// =========================================================
// PARSE TIME
// =========================================================

function parseTime(value) {

    if (
        value instanceof Date
    ) {

        return {

            hour:
                value.getHours(),

            minute:
                value.getMinutes()

        };

    }


    if (
        typeof value === "number"
    ) {

        // Excel time fraction
        if (
            value >= 0 &&
            value < 1
        ) {

            const minutes =
                Math.round(
                    value * 1440
                );


            return {

                hour:
                    Math.floor(
                        minutes / 60
                    ) % 24,

                minute:
                    minutes % 60

            };

        }


        // HHMM
        const number =
            Math.floor(value);


        const hour =
            Math.floor(
                number / 100
            );


        const minute =
            number % 100;


        if (
            hour <= 23 &&
            minute <= 59
        ) {

            return {
                hour,
                minute
            };

        }

    }


    const text =
        String(
            value ?? ""
        ).trim();


    if (!text) {
        return null;
    }


    let match =
        text.match(
            /^(\d{1,2}):(\d{2})/
        );


    if (match) {

        return {

            hour:
                Number(match[1]),

            minute:
                Number(match[2])

        };

    }


    match =
        text.match(
            /^(\d{3,4})$/
        );


    if (match) {

        const number =
            Number(match[1]);


        const hour =
            Math.floor(
                number / 100
            );


        const minute =
            number % 100;


        if (
            hour <= 23 &&
            minute <= 59
        ) {

            return {
                hour,
                minute
            };

        }

    }


    return null;

}


// =========================================================
// BUTTON
// =========================================================

checkBtn.addEventListener(
    "click",
    checkCrew
);


// =========================================================
// ENTER KEY
// =========================================================

departureInput.addEventListener(
    "keydown",
    function (event) {

        if (
            event.key === "Enter"
        ) {

            arrivalInput.focus();

        }

    }
);


arrivalInput.addEventListener(
    "keydown",
    function (event) {

        if (
            event.key === "Enter"
        ) {

            checkCrew();

        }

    }
);


// =========================================================
// MAIN CHECK
// =========================================================

function checkCrew() {

    const fromNumber =
        normalizeFlightNumber(
            departureInput.value
        );


    const toNumber =
        normalizeFlightNumber(
            arrivalInput.value
        );


    // -----------------------------------------------------
    // CHỈ NHẬP CHUYẾN ĐI
    // -----------------------------------------------------

    if (
        fromNumber &&
        !toNumber
    ) {

        showCrewSource(
            fromNumber
        );

        return;

    }


    // -----------------------------------------------------
    // NHẬP CẢ HAI
    // -----------------------------------------------------

    if (
        fromNumber &&
        toNumber
    ) {

        comparePair(
            fromNumber,
            toNumber
        );

        return;

    }


    // -----------------------------------------------------
    // CHỈ NHẬP CHUYẾN ĐẾN
    // -----------------------------------------------------

    if (
        !fromNumber &&
        toNumber
    ) {

        showError(
            "Hãy nhập chuyến bay ở ô TỔ BAY ĐI."
        );

        return;

    }


    showError(
        "Vui lòng nhập số hiệu chuyến bay."
    );

}


// =========================================================
// COMPARE TWO FLIGHTS
// =========================================================
//
// QUAN TRỌNG:
//
// Nếu giống nhau:
//     → KHÔNG ĐỔI TỔ
//
// Nếu khác nhau:
//     → CHỈ HIỆN VJ chuyến ĐI
//
// =========================================================

function comparePair(
    fromNumber,
    toNumber
) {

    const fromFlight =
    findFlight(
        fromNumber,
        {
            fromHAN: true
        }
    );


let toFlight =
    findFlight(
        toNumber
    );


    if (!fromFlight) {

        showError(

            "Không tìm thấy chuyến bay <b>VJ" +
            escapeHtml(
                fromNumber
            ) +
            "</b>."

        );

        return;

    }


    if (!toFlight) {

        showError(

            "Không tìm thấy chuyến bay <b>VJ" +
            escapeHtml(
                toNumber
            ) +
            "</b>."

        );

        return;

    }


    if (
        !fromFlight.crew.length
    ) {

        showError(

            "VJ" +
            escapeHtml(
                fromNumber
            ) +
            " không có dữ liệu tổ bay."

        );

        return;

    }


    if (
        !toFlight.crew.length
    ) {

        showError(

            "VJ" +
            escapeHtml(
                toNumber
            ) +
            " không có dữ liệu tổ bay."

        );

        return;

    }


    // -----------------------------------------------------
    // TẠO SET TÊN
    // -----------------------------------------------------

    const fromSet =
        new Set(
            fromFlight.crew.map(
                crew =>
                    normalizeName(
                        crew.name
                    )
            )
        );


    const toSet =
        new Set(
            toFlight.crew.map(
                crew =>
                    normalizeName(
                        crew.name
                    )
            )
        );


    // -----------------------------------------------------
    // SO SÁNH 2 CHIỀU
    // -----------------------------------------------------

    const removed =
        fromFlight.crew.filter(
            crew =>
                !toSet.has(
                    normalizeName(
                        crew.name
                    )
                )
        );


    const added =
        toFlight.crew.filter(
            crew =>
                !fromSet.has(
                    normalizeName(
                        crew.name
                    )
                )
        );


    // =====================================================
    // KHÔNG ĐỔI TỔ
    // =====================================================

    if (
        removed.length === 0 &&
        added.length === 0
    ) {

        showNoChange(
            fromFlight,
            toFlight
        );

        return;

    }


    // =====================================================
    // CÓ ĐỔI TỔ
    // CHỈ HIỆN CHUYẾN ĐI
    // =====================================================

    showChangedDepartureOnly(
        fromFlight,
        toFlight,
        removed,
        added
    );

}


// =========================================================
// SHOW NO CHANGE
// =========================================================

function showNoChange(
    fromFlight,
    toFlight
) {

    resultBox.innerHTML = `

        <div class="no-change">

            <div class="no-change-icon">
                🟢
            </div>

            <div class="no-change-title">
                KHÔNG ĐỔI TỔ
            </div>

            <div class="no-change-detail">

                VJ${escapeHtml(
                    fromFlight.flight
                )}

                →

                VJ${escapeHtml(
                    toFlight.flight
                )}

                <br>

                Hai chuyến bay có cùng tổ bay.

            </div>

        </div>

    `;

}


// =========================================================
// SHOW CHANGE - ONLY DEPARTURE
// =========================================================

// =========================================================
// SHOW CHANGE - DEPARTURE CREW + CREW SOURCE
// =========================================================

function showChangedDepartureOnly(
    fromFlight,
    toFlight,
    removed,
    added
) {

    // =====================================================
    // TÌM NGUỒN TỪNG THÀNH VIÊN CỦA CHUYẾN ĐI
    // =====================================================

    const crewSources =
        getCrewSources(
            fromFlight
        );


    // =====================================================
    // HIỂN THỊ TỪNG THÀNH VIÊN + NGUỒN
    // =====================================================

    let crewHtml = "";


    crewSources.forEach(
        item => {

            const crew =
                item.crew;


            const source =
                item.source;


            const sourceText =
                source
                    ? "VJ" +
                      source.flight
                    : "HN";


            const sourceClass =
                source
                    ? "crew-source-flight"
                    : "crew-source-hn";


            crewHtml += `

                <div class="crew-member">

                    <div class="crew-role">

                        ${escapeHtml(
                            crew.role
                        )}

                    </div>


                    <div class="crew-name">

                        ${escapeHtml(
                            crew.name
                        )}

                    </div>


                    <div
                        class="${sourceClass}"
                        style="
                            margin-left:auto;
                            min-width:75px;
                            text-align:right;
                            font-size:12px;
                            font-weight:800;
                        "
                    >

                        ← ${escapeHtml(
                            sourceText
                        )}

                    </div>

                </div>

            `;

        }
    );


    // =====================================================
    // TÓM TẮT NGUỒN
    // =====================================================

    const sourceGroups =
        new Map();


    crewSources.forEach(
        item => {

            const key =
                item.source
                    ? "VJ" +
                      item.source.flight
                    : "HN";


            if (
                !sourceGroups.has(key)
            ) {

                sourceGroups.set(
                    key,
                    0
                );

            }


            sourceGroups.set(
                key,
                sourceGroups.get(key) + 1
            );

        }
    );


    let sourceSummary = "";


    sourceGroups.forEach(
        (count, source) => {

            sourceSummary += `

                <span
                    style="
                        display:inline-flex;
                        align-items:center;
                        gap:6px;
                        margin-right:15px;
                        margin-bottom:6px;
                        padding:6px 10px;
                        border-radius:6px;
                        background:${
                            source === "HN"
                                ? "#e2e8f0"
                                : "#dbeafe"
                        };
                        color:${
                            source === "HN"
                                ? "#475569"
                                : "#0f4c81"
                        };
                        font-size:12px;
                        font-weight:800;
                    "
                >

                    ${escapeHtml(
                        source
                    )}

                    <span>
                        ${count} người
                    </span>

                </span>

            `;

        }
    );


    // =====================================================
    // KẾT QUẢ
    // =====================================================

    let html = `

        <div class="change-result">

            <div class="change-icon">
                🔴
            </div>


            <div class="change-title">
                CÓ ĐỔI TỔ
            </div>


            <div class="change-detail">

                Cặp chuyến:

                <b>
                    VJ${escapeHtml(
                        fromFlight.flight
                    )}
                </b>

                →

                <b>
                    VJ${escapeHtml(
                        toFlight.flight
                    )}
                </b>

            </div>

        </div>


        <div class="result-card">

            <div class="result-header">

                <div>

                    <div class="result-title">

                        TỔ BAY VJ${escapeHtml(
                            fromFlight.flight
                        )}

                    </div>


                    <div class="result-subtitle">

                        ${escapeHtml(
                            fromFlight.dep
                        )}

                        →

                        ${escapeHtml(
                            fromFlight.arr
                        )}

                    </div>

                </div>

            </div>


            <!-- =========================================
                 NGUỒN TỔ BAY
                 ========================================= -->

            <div
                style="
                    padding:14px 16px;
                    background:#f8fafc;
                    border-bottom:1px solid #dbe3ec;
                "
            >

                <div
                    style="
                        font-size:11px;
                        font-weight:800;
                        color:#64748b;
                        letter-spacing:.8px;
                        margin-bottom:8px;
                    "
                >

                    NGUỒN TỔ BAY

                </div>


                <div>

                    ${sourceSummary}

                </div>

            </div>


            <!-- =========================================
                 CREW
                 ========================================= -->

            <div class="crew-list">

                ${crewHtml}

            </div>

        </div>

    `;


    // =====================================================
    // CHI TIẾT THAY ĐỔI
    // =====================================================

    if (
        removed.length ||
        added.length
    ) {

        html += `

            <div class="difference">

                <div class="difference-title">

                    THÀNH VIÊN THAY ĐỔI SO VỚI VJ${escapeHtml(
                        toFlight.flight
                    )}

                </div>

        `;


        removed.forEach(
            crew => {

                html += `

                    <div class="difference-row">

                        <span class="removed">
                            ❌
                        </span>

                        <span>
                            ${escapeHtml(
                                crew.role
                            )}
                        </span>

                        <span class="removed">

                            ${escapeHtml(
                                crew.name
                            )}

                        </span>

                    </div>

                `;

            }
        );


        added.forEach(
            crew => {

                html += `

                    <div class="difference-row">

                        <span class="added">
                            ➕
                        </span>

                        <span>
                            ${escapeHtml(
                                crew.role
                            )}
                        </span>

                        <span class="added">

                            ${escapeHtml(
                                crew.name
                            )}

                        </span>

                    </div>

                `;

            }
        );


        html += `

            </div>

        `;

    }


    resultBox.innerHTML =
        html;

} {

    let html = `

        <div class="change-result">

            <div class="change-icon">
                🔴
            </div>

            <div class="change-title">
                CÓ ĐỔI TỔ
            </div>

            <div class="change-detail">

                Cặp chuyến:

                <b>
                    VJ${escapeHtml(
                        fromFlight.flight
                    )}
                </b>

                →

                <b>
                    VJ${escapeHtml(
                        toFlight.flight
                    )}
                </b>

            </div>

        </div>


        <div class="result-card">

            <div class="result-header">

                <div>

                    <div class="result-title">

                        TỔ BAY VJ${escapeHtml(
                            fromFlight.flight
                        )}

                    </div>

                    <div class="result-subtitle">

                        ${escapeHtml(
                            fromFlight.dep
                        )}

                        →

                        ${escapeHtml(
                            fromFlight.arr
                        )}

                    </div>

                </div>

            </div>


            <div class="crew-list">

                ${renderCrewList(
                    fromFlight.crew,
                    removed
                )}

            </div>

        </div>

    `;


    // -----------------------------------------------------
    // CHI TIẾT KHÁC BIỆT
    // -----------------------------------------------------

    if (
        removed.length ||
        added.length
    ) {

        html += `

            <div class="difference">

                <div class="difference-title">

                    THÀNH VIÊN THAY ĐỔI

                </div>

        `;


        removed.forEach(
            crew => {

                html += `

                    <div class="difference-row">

                        <span class="removed">
                            ❌
                        </span>

                        <span>
                            ${escapeHtml(
                                crew.role
                            )}
                        </span>

                        <span class="removed">

                            ${escapeHtml(
                                crew.name
                            )}

                        </span>

                    </div>

                `;

            }
        );


        added.forEach(
            crew => {

                html += `

                    <div class="difference-row">

                        <span class="added">
                            ➕
                        </span>

                        <span>
                            ${escapeHtml(
                                crew.role
                            )}
                        </span>

                        <span class="added">

                            ${escapeHtml(
                                crew.name
                            )}

                        </span>

                    </div>

                `;

            }
        );


        html += `

            </div>

        `;

    }


    resultBox.innerHTML =
        html;

}


// =========================================================
// FIND SOURCE FOR ONE CREW MEMBER
// =========================================================

function findSourceForCrew(
    currentFlight,
    crew
) {

    const currentTime =
        getFlightDateTime(
            currentFlight
        );


    const currentDep =
        normalizeAirport(
            currentFlight.dep
        );


    const personName =
        normalizeName(
            crew.name
        );


    const candidates = [];


    flights.forEach(
        flight => {

            if (
                flight === currentFlight
            ) {

                return;

            }


            if (
                !flight.crew ||
                !flight.crew.length
            ) {

                return;

            }


            // ------------------------------------------------
            // ARR của chuyến trước phải bằng DEP hiện tại
            // ------------------------------------------------

            const arrivalAirport =
                normalizeAirport(
                    flight.arr
                );


            if (
                !currentDep ||
                arrivalAirport !== currentDep
            ) {

                return;

            }


            // ------------------------------------------------
            // Phải xảy ra trước
            // ------------------------------------------------

            const previousTime =
                getFlightDateTime(
                    flight
                );


            if (
                !isFinite(
                    currentTime
                ) ||
                !isFinite(
                    previousTime
                )
            ) {

                return;

            }


            if (
                previousTime >= currentTime
            ) {

                return;

            }


            // ------------------------------------------------
            // Phải có người đó
            // ------------------------------------------------

            const found =
                flight.crew.some(
                    member =>
                        normalizeName(
                            member.name
                        ) === personName
                );


            if (!found) {

                return;

            }


            candidates.push({

                flight:
                    flight,

                difference:
                    currentTime -
                    previousTime

            });

        }
    );


    candidates.sort(
        (a, b) =>
            a.difference -
            b.difference
    );


    if (
        candidates.length
    ) {

        return candidates[0].flight;

    }


    return null;

}


// =========================================================
// GET SOURCE FOR EVERY CREW MEMBER
// =========================================================

function getCrewSources(
    flight
) {

    return flight.crew.map(
        crew => {

            return {

                crew:

                    crew,

                source:

                    findSourceForCrew(
                        flight,
                        crew
                    )

            };

        }
    );

}


// =========================================================
// SHOW CREW SOURCE
// =========================================================

function showCrewSource(
    flightNumber
) {

    const flight =
    findFlight(
        flightNumber,
        {
            fromHAN: true
        }
    );


    if (!flight) {

        showError(

            "Không tìm thấy chuyến bay <b>VJ" +
            escapeHtml(
                flightNumber
            ) +
            "</b>."

        );

        return;

    }


    if (
        !flight.crew.length
    ) {

        showError(

            "VJ" +
            escapeHtml(
                flightNumber
            ) +
            " không có dữ liệu tổ bay."

        );

        return;

    }


    const sources =
        getCrewSources(
            flight
        );


    renderCrewSourceResult(
        flight,
        sources
    );

}


// =========================================================
// RENDER SOURCE RESULT
// =========================================================

function renderCrewSourceResult(
    flight,
    sources
) {

    const groups =
        new Map();


    sources.forEach(
        item => {

            const key =
                item.source
                    ? normalizeFlightNumber(
                        item.source.flight
                    )
                    : "HN";


            if (
                !groups.has(key)
            ) {

                groups.set(
                    key,
                    {

                        source:
                            item.source,

                        members: []

                    }
                );

            }


            groups
                .get(key)
                .members
                .push(
                    item.crew
                );

        }
    );


    // -----------------------------------------------------
    // HN FIRST
    // -----------------------------------------------------

    const sorted =
        Array.from(
            groups.values()
        )
        .sort(
            (a, b) => {

                if (
                    !a.source &&
                    b.source
                ) {

                    return -1;

                }


                if (
                    a.source &&
                    !b.source
                ) {

                    return 1;

                }


                if (
                    !a.source &&
                    !b.source
                ) {

                    return 0;

                }


                return (
                    getFlightDateTime(
                        a.source
                    ) -
                    getFlightDateTime(
                        b.source
                    )
                );

            }
        );


    let summary = "";


    sorted.forEach(
        group => {

            const sourceName =
                group.source
                    ? "VJ" +
                      group.source.flight
                    : "HN";


            const roles =
                group.members
                    .map(
                        crew =>
                            crew.role ||
                            "CREW"
                    )
                    .join(", ");


            summary += `

                <div
                    style="
                        display:flex;
                        align-items:center;
                        gap:15px;
                        padding:10px 0;
                        border-bottom:1px solid #e2e8f0;
                    "
                >

                    <strong
                        style="
                            min-width:70px;
                            color:#0f4c81;
                        "
                    >

                        ${escapeHtml(
                            sourceName
                        )}

                    </strong>

                    <span>

                        ${escapeHtml(
                            roles
                        )}

                    </span>

                </div>

            `;

        }
    );


    let detail = "";


    sources.forEach(
        item => {

            const sourceName =
                item.source
                    ? "VJ" +
                      item.source.flight
                    : "HN";


            detail += `

                <div class="crew-member">

                    <div class="crew-role">

                        ${escapeHtml(
                            item.crew.role
                        )}

                    </div>

                    <div class="crew-name">

                        ${escapeHtml(
                            item.crew.name
                        )}

                    </div>

                    <div
                        style="
                            margin-left:auto;
                            font-size:12px;
                            font-weight:800;
                            color:${
                                item.source
                                    ? "#0f4c81"
                                    : "#64748b"
                            };
                        "
                    >

                        ← ${escapeHtml(
                            sourceName
                        )}

                    </div>

                </div>

            `;

        }
    );


    resultBox.innerHTML = `

        <div class="result-card">

            <div class="result-header">

                <div>

                    <div class="result-title">

                        VJ${escapeHtml(
                            flight.flight
                        )}

                    </div>

                    <div class="result-subtitle">

                        ${escapeHtml(
                            flight.dep
                        )}

                        →

                        ${escapeHtml(
                            flight.arr
                        )}

                    </div>

                </div>

            </div>


            <div
                style="
                    padding:20px;
                    background:#f8fafc;
                    border-bottom:1px solid #dbe3ec;
                "
            >

                <div
                    style="
                        font-size:12px;
                        color:#64748b;
                        font-weight:800;
                        letter-spacing:1px;
                        margin-bottom:10px;
                    "
                >

                    TỔ BAY ĐẾN TỪ

                </div>


                ${summary}

            </div>


            <div class="crew-list">

                ${detail}

            </div>

        </div>

    `;

}


// =========================================================
// RENDER CREW LIST
// =========================================================

function renderCrewList(
    crewList,
    changedCrew
) {

    const changedNames =
        new Set(
            changedCrew.map(
                crew =>
                    normalizeName(
                        crew.name
                    )
            )
        );


    return crewList
        .map(
            crew => {

                const isChanged =
                    changedNames.has(
                        normalizeName(
                            crew.name
                        )
                    );


                return `

                    <div
                        class="crew-member"
                        style="${
                            isChanged
                                ? "background:#fff7ed;"
                                : ""
                        }"
                    >

                        <div class="crew-role">

                            ${escapeHtml(
                                crew.role
                            )}

                        </div>

                        <div class="crew-name">

                            ${escapeHtml(
                                crew.name
                            )}

                        </div>

                    </div>

                `;

            }
        )
        .join("");

}


// =========================================================
// ERROR
// =========================================================

function showError(
    message
) {

    resultBox.innerHTML = `

        <div class="notfound">

            ❌ ${message}

        </div>

    `;

}


// =========================================================
// COLUMN LETTER
// =========================================================

function columnLetter(
    index
) {

    if (
        index < 0
    ) {

        return "?";

    }


    let result = "";
    let number = index + 1;


    while (
        number > 0
    ) {

        const remainder =
            (number - 1) % 26;


        result =
            String.fromCharCode(
                65 + remainder
            ) +
            result;


        number =
            Math.floor(
                (number - 1) / 26
            );

    }


    return result;

}


// =========================================================
// ESCAPE HTML
// =========================================================

function escapeHtml(
    text
) {

    return String(
        text ?? ""
    )
    .replace(
        /&/g,
        "&amp;"
    )
    .replace(
        /</g,
        "&lt;"
    )
    .replace(
        />/g,
        "&gt;"
    )
    .replace(
        /"/g,
        "&quot;"
    )
    .replace(
        /'/g,
        "&#039;"
    );

}
