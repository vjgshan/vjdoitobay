let flights = [];


// =========================================================
// GET ELEMENTS
// =========================================================

const excelFile =
    document.getElementById("excelFile");

const fileName =
    document.getElementById("fileName");

const statusBox =
    document.getElementById("status");

const departureInput =
    document.getElementById("departureFlight");

const arrivalInput =
    document.getElementById("arrivalFlight");

const checkBtn =
    document.getElementById("checkBtn");

const resultBox =
    document.getElementById("result");




if (excelFile) {

    excelFile.addEventListener(
        "change",
        function (event) {

            const file =
                event.target.files[0];

            if (!file) {
                return;
            }

            fileName.textContent =
                file.name;

            statusBox.innerHTML =
                "⏳ Đang đọc Crew List...";


            const reader =
                new FileReader();


            reader.onload =
                function (e) {

                    try {

                        const workbook =
                            XLSX.read(
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

                        console.error(
                            "Excel error:",
                            error
                        );

                        flights = [];

                        statusBox.innerHTML =
                            "❌ Không thể đọc file Excel.";

                    }

                };


            reader.readAsArrayBuffer(file);

        }
    );

}



function normalizeHeader(value) {

    return String(
        value ?? ""
    )
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ")
    .replace(/\n/g, " ");

}




function findColumn(
    header,
    names
) {

    for (
        let i = 0;
        i < header.length;
        i++
    ) {

        const value =
            normalizeHeader(
                header[i]
            );


        if (!value) {
            continue;
        }


        for (
            const name of names
        ) {

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


        for (
            const cell of row
        ) {

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



function parseCrewList(rows) {

    flights = [];


    const headerRow =
        findHeaderRow(rows);


    if (
        headerRow === -1
    ) {

        statusBox.innerHTML =
            "❌ Không tìm thấy dòng header FLT / CREW.";

        return;

    }


    const header =
        rows[headerRow];


    // -----------------------------------------------------
    // COLUMNS
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


    if (
        flightCol === -1
    ) {

        statusBox.innerHTML =
            "❌ Không tìm thấy cột FLT.";

        return;

    }


    if (
        crewCol === -1
    ) {

        statusBox.innerHTML =
            "❌ Không tìm thấy cột CREW.";

        return;

    }


    console.log(
        "================================"
    );

    console.log(
        "HEADER ROW:",
        headerRow + 1
    );

    console.log(
        "FLT:",
        columnLetter(flightCol)
    );

    console.log(
        "CREW:",
        columnLetter(crewCol)
    );

    console.log(
        "CREW #:",
        crewCountCol >= 0
            ? columnLetter(crewCountCol)
            : "Không có"
    );




    let currentFlight = null;


    for (
        let i = headerRow + 1;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i] || [];



        const rawFlight =
            row[flightCol];


        const flightText =
            String(
                rawFlight ?? ""
            ).trim();


        if (
            flightText !== ""
        ) {

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


        if (
            !currentFlight
        ) {

            continue;

        }


        const crewValue =
            row[crewCol];


        if (
            crewValue === undefined ||
            crewValue === null ||
            String(
                crewValue
            ).trim() === ""
        ) {

            continue;

        }


        const members =
            String(
                crewValue
            )
            .split(/\r?\n/)
            .map(
                x =>
                    x.trim()
            )
            .filter(
                x =>
                    x !== ""
            );


        members.forEach(
            member => {

                const crew =
                    parseCrew(member);


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
        (
            headerRow + 1
        ) +

        " | FLT: cột " +
        columnLetter(flightCol) +

        " | CREW: cột " +
        columnLetter(crewCol) +

        (
            crewCountCol >= 0
                ? " | CREW #: cột " +
                  columnLetter(crewCountCol)
                : ""
        ) +

        "</small>";


    console.log(
        "FLIGHTS:",
        flights
    );

}



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




function normalizeAirport(value) {

    return String(
        value ?? ""
    )
    .trim()
    .toUpperCase();

}




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




function parseCrew(text) {

    let value =
        String(
            text ?? ""
        ).trim();


    let role = "";


    

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
                        x =>
                            x.trim()
                    )
                    .filter(
                        x =>
                            x &&
                            x.toUpperCase() !==
                            "DHD"
                    )
                    .join(
                        " / "
                    );


                if (inside) {

                    roles.push(
                        inside
                    );

                }

            }
        );


        role =
            roles.join(
                " / "
            );

    }




    value =
        value.replace(
            /\(.*?\)/g,
            ""
        );


   

    value =
        value.replace(
            /^\s*-\s*/,
            ""
        );




    value =
        value.replace(
            /\+/g,
            ""
        );


    value =
        value.trim();


    return {

        role:
            role,

        name:
            value

    };

}




function findFlight(
    number,
    options = {}
) {

    const target =
        normalizeFlightNumber(
            number
        );


    const candidates =
        flights.filter(
            flight =>

                normalizeFlightNumber(
                    flight.flight
                ) === target

        );


    if (
        !candidates.length
    ) {

        return null;

    }


  

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



    if (
        candidates.length === 1
    ) {

        return candidates[0];

    }




    return candidates[0];

}


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
                value *
                86400000
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

        // Excel time
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




function getFlightDateTime(flight) {

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


    const date =
        parseDate(
            flight.date
        );


    const parsedTime =
        parseTime(
            time
        );


    if (
        !date ||
        !parsedTime
    ) {

        return NaN;

    }


    return new Date(

        date.year,
        date.month - 1,
        date.day,

        parsedTime.hour,
        parsedTime.minute,
        0,
        0

    ).getTime();

}




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
                flight ===
                currentFlight
            ) {

                return;

            }


            if (
                !flight.crew ||
                !flight.crew.length
            ) {

                return;

            }


         

            const previousArr =
                normalizeAirport(
                    flight.arr
                );


            if (
                !currentDep ||
                previousArr !== currentDep
            ) {

                return;

            }


           

            const previousTime =
                getFlightDateTime(
                    flight
                );


            if (
                !isFinite(currentTime) ||
                !isFinite(previousTime)
            ) {

                return;

            }


            if (
                previousTime >=
                currentTime
            ) {

                return;

            }


            

            const found =
                flight.crew.some(
                    member =>

                        normalizeName(
                            member.name
                        ) ===
                        personName
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




function getCrewSources(flight) {

    return flight.crew.map(
        crew => ({

            crew:
                crew,

            source:
                findSourceForCrew(
                    flight,
                    crew
                )

        })
    );

}




function checkCrew() {

    const fromNumber =
        normalizeFlightNumber(
            departureInput
                ? departureInput.value
                : ""
        );


    const toNumber =
        normalizeFlightNumber(
            arrivalInput
                ? arrivalInput.value
                : ""
        );


    

    if (
        fromNumber &&
        !toNumber
    ) {

        showCrewSource(
            fromNumber
        );

        return;

    }


  

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


    

    const toFlight =
        findFlight(
            toNumber
        );


    if (!fromFlight) {

        showError(

            "Không tìm thấy chuyến bay <b>VJ" +
            escapeHtml(fromNumber) +
            "</b>."

        );

        return;

    }


    if (!toFlight) {

        showError(

            "Không tìm thấy chuyến bay <b>VJ" +
            escapeHtml(toNumber) +
            "</b>."

        );

        return;

    }


    

    if (
        !fromFlight.crew ||
        !fromFlight.crew.length
    ) {

        showError(

            "VJ" +
            escapeHtml(fromNumber) +
            " không có dữ liệu tổ bay."

        );

        return;

    }


    if (
        !toFlight.crew ||
        !toFlight.crew.length
    ) {

        showError(

            "VJ" +
            escapeHtml(toNumber) +
            " không có dữ liệu tổ bay."

        );

        return;

    }


    

    const toSet =
        new Set(

            toFlight.crew.map(
                crew =>
                    normalizeName(
                        crew.name
                    )
            )

        );


    

    const keptCrew =
        fromFlight.crew.filter(

            crew =>
                toSet.has(
                    normalizeName(
                        crew.name
                    )
                )

        );


    

    const notKeptCrew =
        fromFlight.crew.filter(

            crew =>
                !toSet.has(
                    normalizeName(
                        crew.name
                    )
                )

        );


    const totalCrew =
        fromFlight.crew.length;


    const keptCount =
        keptCrew.length;


    const notKeptCount =
        notKeptCrew.length;


    const keptPercent =
        totalCrew > 0

            ? (
                keptCount /
                totalCrew
            ) * 100

            : 0;


    console.log(
        "=============================="
    );

    console.log(
        "TỔ BAY ĐI:",
        fromFlight.flight
    );

    console.log(
        "TỔ BAY ĐẾN:",
        toFlight.flight
    );

    console.log(
        "TỔNG:",
        totalCrew
    );

    console.log(
        "GIỮ:",
        keptCount
    );

    console.log(
        "KHÔNG GIỮ:",
        notKeptCount
    );

    console.log(
        "TỶ LỆ:",
        keptPercent + "%"
    );


  

    if (
        keptPercent > 50
    ) {

        showNoChangeMajority(
            fromFlight,
            toFlight,
            keptCount,
            totalCrew,
            notKeptCrew
        );

        return;

    }


    

    showChangedDepartureOnly(
        fromFlight,
        toFlight,
        notKeptCrew
    );

}




function showNoChangeMajority(
    fromFlight,
    toFlight,
    keptCount,
    totalCrew,
    notKeptCrew
) {

    const percent =
        totalCrew > 0
            ? Math.round(
                (
                    keptCount /
                    totalCrew
                ) * 100
            )
            : 0;


   

    const notKeptSources =
        notKeptCrew.map(
            crew => {

                return {

                    crew: crew,

                    source:
                        findSourceForCrew(
                            fromFlight,
                            crew
                        )

                };

            }
        );


   

    let membersHtml = "";


    notKeptSources.forEach(
        item => {

            const source =
                item.source
                    ? "VJ" +
                      item.source.flight
                    : "HN";


            membersHtml += `

                <div
                    style="
                        display:flex;
                        align-items:center;
                        gap:10px;
                        padding:11px 14px;
                        border-bottom:1px solid #e5e7eb;
                        background:#ffffff;
                    "
                >

                    <!-- ROLE -->

                    <div
                        style="
                            min-width:42px;
                            color:#64748b;
                            font-size:12px;
                            font-weight:800;
                        "
                    >

                        ${escapeHtml(
                            item.crew.role
                        )}

                    </div>


                    <!-- NAME -->

                    <div
                        style="
                            flex:1;
                            color:#1e3a5f;
                            font-size:13px;
                            font-weight:700;
                        "
                    >

                        ${escapeHtml(
                            item.crew.name
                        )}

                    </div>


                    <!-- SOURCE -->

                    <div
                        style="
                            min-width:65px;
                            text-align:right;
                            color:${
                                source === "HN"
                                    ? "#64748b"
                                    : "#e21b23"
                            };
                            font-size:12px;
                            font-weight:800;
                        "
                    >

                        ← ${escapeHtml(
                            source
                        )}

                    </div>

                </div>

            `;

        }
    );


   

    let notKeptHtml = "";


    if (
        notKeptSources.length > 0
    ) {

        notKeptHtml = `

            <div
                style="
                    margin-top:14px;
                    background:#f8fafc;
                    border:1px solid #dbe3ec;
                    border-radius:10px;
                    overflow:hidden;
                "
            >

                <div
                    style="
                        padding:12px 14px;
                        background:#eef2f7;
                        color:#334155;
                        font-size:11px;
                        font-weight:800;
                        letter-spacing:.6px;
                    "
                >

                    THÀNH VIÊN KHÔNG TIẾP TỤC
                    Ở VJ${escapeHtml(
                        toFlight.flight
                    )}

                </div>


                ${membersHtml}

            </div>

        `;

    }



    resultBox.innerHTML = `

        <div class="no-change">

            <div class="no-change-icon">
                ✓
            </div>


            <div class="no-change-title">
                KHÔNG ĐỔI TỔ
            </div>


            <div class="no-change-detail">

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


                <br><br>


                <strong>
                    ${keptCount}/${totalCrew}
                </strong>

                thành viên của tổ bay đi
                vẫn tiếp tục ở

                <strong>
                    VJ${escapeHtml(
                        toFlight.flight
                    )}
                </strong>.


                <br>


                <strong>
                    ${percent}% tổ bay được giữ nguyên.
                </strong>


                ${notKeptHtml}

            </div>

        </div>

    `;

}




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
            escapeHtml(flightNumber) +
            "</b>."

        );

        return;

    }


    if (
        !flight.crew ||
        !flight.crew.length
    ) {

        showError(

            "VJ" +
            escapeHtml(flightNumber) +
            " không có dữ liệu tổ bay."

        );

        return;

    }


    const sources =
        getCrewSources(flight);


    renderCrewSourceResult(
        flight,
        sources
    );

}




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

                    ? "VJ" +
                      item.source.flight

                    : "HN";


            if (
                !groups.has(key)
            ) {

                groups.set(
                    key,
                    {
                        source:
                            item.source,

                        members:
                            []
                    }
                );

            }


            groups
                .get(key)
                .members
                .push(item.crew);

        }
    );


   

    let summary = "";


    groups.forEach(
        (
            group,
            key
        ) => {

            const count =
                group.members.length;


            const isHN =
                key === "HN";


            summary += `

                <span
                    style="
                        display:inline-flex;
                        align-items:center;
                        gap:6px;
                        margin-right:8px;
                        margin-bottom:6px;
                        padding:6px 10px;
                        border-radius:6px;
                        background:${
                            isHN
                                ? "#e2e8f0"
                                : "#dbeafe"
                        };
                        color:${
                            isHN
                                ? "#475569"
                                : "#0f4c81"
                        };
                        font-size:12px;
                        font-weight:800;
                    "
                >

                    ${escapeHtml(key)}

                    <span>
                        ${count} người
                    </span>

                </span>

            `;

        }
    );


 

    let detail = "";


    sources.forEach(
        item => {

            const source =
                item.source

                    ? "VJ" +
                      item.source.flight

                    : "HN";


            detail += `

                <div
                    class="crew-member"
                >

                    <div
                        class="crew-role"
                    >

                        ${escapeHtml(
                            item.crew.role
                        )}

                    </div>


                    <div
                        class="crew-name"
                    >

                        ${escapeHtml(
                            item.crew.name
                        )}

                    </div>


                    <div
                        style="
                            margin-left:auto;
                            min-width:70px;
                            text-align:right;
                            font-size:12px;
                            font-weight:800;
                            color:${
                                source === "HN"
                                    ? "#64748b"
                                    : "#e21b23"
                            };
                        "
                    >

                        ← ${escapeHtml(
                            source
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

                        TỔ BAY VJ${escapeHtml(
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
                    padding:16px 20px;
                    background:#f8fafc;
                    border-bottom:1px solid #dbe3ec;
                "
            >

                <div
                    style="
                        font-size:11px;
                        color:#64748b;
                        font-weight:800;
                        margin-bottom:9px;
                        letter-spacing:.8px;
                    "
                >

                    NGUỒN TỔ BAY

                </div>


                ${summary}

            </div>


            <div class="crew-list">

                ${detail}

            </div>

        </div>

    `;

}




function showChangedDepartureOnly(
    fromFlight,
    toFlight,
    notKeptCrew
) {

    const crewSources =
        getCrewSources(
            fromFlight
        );


    

    let crewHtml = "";


    crewSources.forEach(
        item => {

            const source =
                item.source

                    ? "VJ" +
                      item.source.flight

                    : "HN";


            crewHtml += `

                <div
                    class="crew-member"
                >

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
                            min-width:70px;
                            text-align:right;
                            font-size:12px;
                            font-weight:800;
                            color:${
                                source === "HN"
                                    ? "#64748b"
                                    : "#e21b23"
                            };
                        "
                    >

                        ← ${escapeHtml(
                            source
                        )}

                    </div>

                </div>

            `;

        }
    );


   

    const sourceGroups =
        new Map();


    crewSources.forEach(
        item => {

            const source =
                item.source

                    ? "VJ" +
                      item.source.flight

                    : "HN";


            if (
                !sourceGroups.has(source)
            ) {

                sourceGroups.set(
                    source,
                    0
                );

            }


            sourceGroups.set(
                source,
                sourceGroups.get(source) + 1
            );

        }
    );


    let sourceSummary = "";


    sourceGroups.forEach(
        (
            count,
            source
        ) => {

            const isHN =
                source === "HN";


            sourceSummary += `

                <span
                    style="
                        display:inline-flex;
                        align-items:center;
                        gap:6px;
                        margin-right:8px;
                        margin-bottom:6px;
                        padding:6px 10px;
                        border-radius:6px;
                        background:${
                            isHN
                                ? "#e2e8f0"
                                : "#dbeafe"
                        };
                        color:${
                            isHN
                                ? "#475569"
                                : "#0f4c81"
                        };
                        font-size:12px;
                        font-weight:800;
                    "
                >

                    ${escapeHtml(source)}

                    <span>
                        ${count} người
                    </span>

                </span>

            `;

        }
    );


    

    let html = `

        <div class="change-result">

            <div class="change-icon">
                !
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
                        color:#64748b;
                        font-weight:800;
                        letter-spacing:.8px;
                        margin-bottom:9px;
                    "
                >

                    NGUỒN TỔ BAY

                </div>


                ${sourceSummary}

            </div>


            <div class="crew-list">

                ${crewHtml}

            </div>

        </div>

    `;


   

    if (
        notKeptCrew &&
        notKeptCrew.length
    ) {

        html += `

            <div class="difference">

                <div class="difference-title">

                    THÀNH VIÊN TỔ BAY ĐI
                    KHÔNG CÒN Ở
                    VJ${escapeHtml(
                        toFlight.flight
                    )}

                </div>

        `;


        notKeptCrew.forEach(
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


        html += `

            </div>

        `;

    }


    resultBox.innerHTML =
        html;

}




function showError(message) {

    resultBox.innerHTML = `

        <div class="notfound">

            ❌ ${message}

        </div>

    `;

}


// =========================================================
// COLUMN LETTER
// =========================================================

function columnLetter(index) {

    if (
        index < 0
    ) {

        return "?";

    }


    let result = "";

    let number =
        index + 1;


    while (
        number > 0
    ) {

        const remainder =
            (
                number - 1
            ) % 26;


        result =
            String.fromCharCode(
                65 + remainder
            ) +
            result;


        number =
            Math.floor(
                (
                    number - 1
                ) / 26
            );

    }


    return result;

}



function escapeHtml(text) {

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




if (
    departureInput
) {

    departureInput.addEventListener(
        "keydown",
        function (event) {

            if (
                event.key === "Enter"
            ) {

                event.preventDefault();


                if (
                    arrivalInput
                ) {

                    arrivalInput.focus();

                }

            }

        }
    );

}


if (
    arrivalInput
) {

    arrivalInput.addEventListener(
        "keydown",
        function (event) {

            if (
                event.key === "Enter"
            ) {

                event.preventDefault();

                checkCrew();

            }

        }
    );

}



if (
    checkBtn
) {

    checkBtn.addEventListener(
        "click",
        checkCrew
    );

}




function onlyNumberInput(input) {

    if (!input) {
        return;
    }


    input.addEventListener(
        "input",
        function () {

            this.value =
                this.value.replace(
                    /[^0-9]/g,
                    ""
                );

        }
    );

}


onlyNumberInput(
    departureInput
);

onlyNumberInput(
    arrivalInput
);



