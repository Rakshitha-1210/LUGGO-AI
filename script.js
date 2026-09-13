/* =========================================================
   LUGGO AI
   AI-Powered Luggage Transportation and Tracking System
   ========================================================= */


/* ================= GLOBAL VARIABLES ================= */

let transportType = "Railway Station";

let allData = [];

let currentBooking = {
    passengerName: "",
    transportType: "",
    station: "",
    currentPoint: "",
    destinationPoint: "",
    bags: 1,
    deliveryTime: "",
    priority: "",
    bookingId: "",
    otp: ""
};


/* =========================================================
   LOAD EXCEL DATA
   ========================================================= */

window.addEventListener("DOMContentLoaded", () => {

    loadExcelData();

    updateTransportUI();

    document
        .getElementById("deliveryTime")
        .addEventListener("change", calculatePriority);

});


async function loadExcelData() {

    try {

        const response = await fetch(
            "Telangana_Transport_Platform_Wise_Data.xlsx"
        );

        if (!response.ok) {
            throw new Error("Excel file could not be loaded.");
        }

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });


        /*
         * IMPORTANT:
         * Your dataset has:
         *
         * Platform-wise Data
         * Summary
         *
         * We use the first sheet.
         */

        const firstSheetName = workbook.SheetNames[0];

        const worksheet = workbook.Sheets[firstSheetName];

        const rawData = XLSX.utils.sheet_to_json(
            worksheet,
            {
                defval: ""
            }
        );


        /*
         * Convert Excel data into clean objects.
         */

        allData = rawData.map(row => {

            return {

                transportType:
                    String(row["Transport Type"] || "").trim(),

                station:
                    String(
                        row["Station / Terminal Name"] || ""
                    ).trim(),

                district:
                    String(row["District"] || "").trim(),

                point:
                    String(
                        row["Platform / Terminal"] || ""
                    ).trim(),

                number:
                    row["No."]

            };

        }).filter(row => {

            return (
                row.transportType !== "" &&
                row.station !== "" &&
                row.point !== ""
            );

        });


        console.log(
            "LUGGO Excel Data Loaded:",
            allData
        );


        populateStations();

    }

    catch (error) {

        console.error(error);

        showError(
            "Unable to load Excel dataset. Make sure the Excel file is in the same folder as index.html."
        );

    }

}


/* =========================================================
   TRANSPORT SELECTION
   ========================================================= */

function selectTransport(type) {

    transportType = type;

    updateTransportUI();

    populateStations();

}


/* =========================================================
   UPDATE UI FOR RAILWAY / BUS
   ========================================================= */

function updateTransportUI() {

    const railwayCard =
        document.getElementById("railwayCard");

    const busCard =
        document.getElementById("busCard");


    const stationLabel =
        document.getElementById("stationLabel");

    const currentLabel =
        document.getElementById("currentLabel");

    const destinationLabel =
        document.getElementById("destinationLabel");


    if (transportType === "Railway Station") {

        railwayCard.classList.add("active");

        busCard.classList.remove("active");


        stationLabel.textContent =
            "Railway Station";

        currentLabel.textContent =
            "Current Platform";

        destinationLabel.textContent =
            "Destination Platform";


    } else {

        railwayCard.classList.remove("active");

        busCard.classList.add("active");


        stationLabel.textContent =
            "Bus Station";

        currentLabel.textContent =
            "Current Terminal";

        destinationLabel.textContent =
            "Destination Terminal";

    }


    clearPointDropdowns();

}


/* =========================================================
   POPULATE STATIONS
   ========================================================= */

function populateStations() {

    const stationSelect =
        document.getElementById("stationSelect");


    stationSelect.innerHTML = "";


    const defaultOption =
        document.createElement("option");

    defaultOption.value = "";

    defaultOption.textContent =
        transportType === "Railway Station"
            ? "Select Railway Station"
            : "Select Bus Station";

    stationSelect.appendChild(defaultOption);


    /*
     * Get stations for selected transport type.
     */

    const stations = allData

        .filter(row => {

            return (
                row.transportType === transportType
            );

        })

        .map(row => row.station);


    /*
     * Remove duplicate station names.
     */

    const uniqueStations =
        [...new Set(stations)];


    uniqueStations.sort();


    uniqueStations.forEach(station => {

        const option =
            document.createElement("option");

        option.value = station;

        option.textContent = station;

        stationSelect.appendChild(option);

    });


    /*
     * When station changes,
     * populate platforms / terminals.
     */

    stationSelect.onchange =
        populatePoints;


    clearPointDropdowns();

}


/* =========================================================
   POPULATE PLATFORM / TERMINAL
   ========================================================= */

function populatePoints() {

    const station =
        document.getElementById("stationSelect").value;


    const currentSelect =
        document.getElementById("currentSelect");

    const destinationSelect =
        document.getElementById("destinationSelect");


    currentSelect.innerHTML = "";

    destinationSelect.innerHTML = "";


    const currentDefault =
        document.createElement("option");

    currentDefault.value = "";

    currentDefault.textContent =
        transportType === "Railway Station"
            ? "Select Current Platform"
            : "Select Current Terminal";

    currentSelect.appendChild(currentDefault);


    const destinationDefault =
        document.createElement("option");

    destinationDefault.value = "";

    destinationDefault.textContent =
        transportType === "Railway Station"
            ? "Select Destination Platform"
            : "Select Destination Terminal";

    destinationSelect.appendChild(
        destinationDefault
    );


    if (!station) {
        return;
    }


    /*
     * Get platform / terminal data
     * for selected station.
     */

    const points = allData

        .filter(row => {

            return (
                row.transportType === transportType &&
                row.station === station
            );

        })

        .map(row => row.point);


    const uniquePoints =
        [...new Set(points)];


    /*
     * Sort Platform 1, Platform 2...
     * and Terminal 1, Terminal 2...
     */

    uniquePoints.sort((a, b) => {

        const numberA =
            parseInt(a.match(/\d+/)?.[0] || 0);

        const numberB =
            parseInt(b.match(/\d+/)?.[0] || 0);

        return numberA - numberB;

    });


    uniquePoints.forEach(point => {

        const currentOption =
            document.createElement("option");

        currentOption.value = point;

        currentOption.textContent = point;

        currentSelect.appendChild(
            currentOption
        );


        const destinationOption =
            document.createElement("option");

        destinationOption.value = point;

        destinationOption.textContent = point;

        destinationSelect.appendChild(
            destinationOption
        );

    });

}


/* =========================================================
   CLEAR PLATFORM / TERMINAL DROPDOWNS
   ========================================================= */

function clearPointDropdowns() {

    const currentSelect =
        document.getElementById("currentSelect");

    const destinationSelect =
        document.getElementById("destinationSelect");


    currentSelect.innerHTML = "";

    destinationSelect.innerHTML = "";


    const currentOption =
        document.createElement("option");

    currentOption.value = "";

    currentOption.textContent =
        transportType === "Railway Station"
            ? "Select Current Platform"
            : "Select Current Terminal";


    currentSelect.appendChild(
        currentOption
    );


    const destinationOption =
        document.createElement("option");

    destinationOption.value = "";

    destinationOption.textContent =
        transportType === "Railway Station"
            ? "Select Destination Platform"
            : "Select Destination Terminal";


    destinationSelect.appendChild(
        destinationOption
    );

}


/* =========================================================
   AI PRIORITY
   ========================================================= */

function calculatePriority() {

    const deliveryTime =
        parseInt(
            document.getElementById("deliveryTime").value
        );


    const priorityText =
        document.getElementById("priorityText");


    if (!deliveryTime) {

        priorityText.textContent =
            "Select required delivery time to calculate priority";

        return "";

    }


    let priority = "";


    if (deliveryTime <= 5) {

        priority = "CRITICAL";

    }

    else if (deliveryTime <= 10) {

        priority = "HIGH";

    }

    else if (deliveryTime <= 20) {

        priority = "MEDIUM";

    }

    else {

        priority = "NORMAL";

    }


    priorityText.textContent =
        `AI Priority: ${priority}`;


    return priority;

}


/* =========================================================
   BOOK LUGGAGE
   ========================================================= */

function bookLuggage() {

    clearError();


    const passengerName =
        document.getElementById("passengerName")
            .value
            .trim();


    const station =
        document.getElementById("stationSelect")
            .value;


    const currentPoint =
        document.getElementById("currentSelect")
            .value;


    const destinationPoint =
        document.getElementById("destinationSelect")
            .value;


    const bags =
        document.getElementById("bags")
            .value;


    const deliveryTime =
        document.getElementById("deliveryTime")
            .value;


    /* ================= VALIDATION ================= */


    if (!passengerName) {

        showError(
            "Please enter passenger name."
        );

        return;

    }


    if (!station) {

        showError(
            "Please select a station or bus station."
        );

        return;

    }


    if (!currentPoint) {

        showError(
            transportType === "Railway Station"
                ? "Please select current platform."
                : "Please select current terminal."
        );

        return;

    }


    if (!destinationPoint) {

        showError(
            transportType === "Railway Station"
                ? "Please select destination platform."
                : "Please select destination terminal."
        );

        return;

    }


    if (currentPoint === destinationPoint) {

        showError(
            "Current and destination cannot be the same."
        );

        return;

    }


    if (!bags || bags < 1) {

        showError(
            "Please enter the number of bags."
        );

        return;

    }


    if (!deliveryTime) {

        showError(
            "Please select required delivery time."
        );

        return;

    }


    /* ================= AI PRIORITY ================= */

    const priority =
        calculatePriority();


    /* ================= BOOKING ID ================= */

    const bookingId =
        "LUGGO" +
        Math.floor(
            100000 +
            Math.random() * 900000
        );


    /* ================= OTP ================= */

    const otp =
        Math.floor(
            100000 +
            Math.random() * 900000
        ).toString();


    /* ================= SAVE BOOKING ================= */

    currentBooking = {

        passengerName:
            passengerName,

        transportType:
            transportType,

        station:
            station,

        currentPoint:
            currentPoint,

        destinationPoint:
            destinationPoint,

        bags:
            bags,

        deliveryTime:
            deliveryTime,

        priority:
            priority,

        bookingId:
            bookingId,

        otp:
            otp

    };


    /* ================= SHOW CONFIRMATION ================= */

    document.getElementById(
        "confirmPassenger"
    ).textContent =
        passengerName;


    document.getElementById(
        "bookingId"
    ).textContent =
        bookingId;


    document.getElementById(
        "confirmTransport"
    ).textContent =
        transportType;


    document.getElementById(
        "confirmStation"
    ).textContent =
        station;


    document.getElementById(
        "confirmFrom"
    ).textContent =
        currentPoint;


    document.getElementById(
        "confirmTo"
    ).textContent =
        destinationPoint;


    document.getElementById(
        "confirmBags"
    ).textContent =
        bags;


    document.getElementById(
        "confirmPriority"
    ).textContent =
        priority;


    document.getElementById(
        "otpValue"
    ).textContent =
        otp;


    /* ================= QR ================= */

    const qrContainer =
        document.getElementById("qrcode");


    qrContainer.innerHTML = "";


    const qrData =

        `LUGGO AI
Booking ID: ${bookingId}
Passenger: ${passengerName}
Transport: ${transportType}
Station: ${station}
From: ${currentPoint}
To: ${destinationPoint}
Bags: ${bags}`;


    new QRCode(
        qrContainer,
        {
            text: qrData,
            width: 180,
            height: 180
        }
    );


    /* ================= DISPLAY ================= */

    document
        .getElementById("confirmation")
        .classList.remove("hidden");


    document
        .getElementById("confirmation")
        .scrollIntoView({
            behavior: "smooth"
        });


}


/* =========================================================
   SHOW TRACKING
   ========================================================= */

function showTracking() {

    document
        .getElementById("tracking")
        .classList.remove("hidden");


    document
        .getElementById("trackingBookingId")
        .textContent =
        currentBooking.bookingId;


    document
        .getElementById("tracking")
        .scrollIntoView({
            behavior: "smooth"
        });


    startTracking();

}


/* =========================================================
   TRACKING SIMULATION
   ========================================================= */

function startTracking() {

    resetTracking();


    /*
     * Step 1 - Booked
     */

    updateTrackingStep(
        1,
        "BOOKED"
    );


    /*
     * Step 2 - Picked Up
     */

    setTimeout(() => {

        updateTrackingStep(
            2,
            "PICKED UP"
        );

    }, 3000);


    /*
     * Step 3 - In Transit
     */

    setTimeout(() => {

        updateTrackingStep(
            3,
            "IN TRANSIT"
        );

    }, 6000);


    /*
     * Step 4 - Destination Reached
     */

    setTimeout(() => {

        updateTrackingStep(
            4,
            "DESTINATION REACHED"
        );


        document
            .getElementById("verification")
            .classList.remove("hidden");


        document
            .getElementById("verification")
            .scrollIntoView({
                behavior: "smooth"
            });


    }, 9000);

}


/* =========================================================
   UPDATE TRACKING STEP
   ========================================================= */

function updateTrackingStep(
    stepNumber,
    status
) {

    for (
        let i = 1;
        i <= stepNumber;
        i++
    ) {

        const step =
            document.getElementById(
                `step${i}`
            );

        step.classList.add(
            "completed"
        );

    }


    document
        .getElementById("currentStatus")
        .textContent =
        status;

}


/* =========================================================
   RESET TRACKING
   ========================================================= */

function resetTracking() {

    for (
        let i = 1;
        i <= 5;
        i++
    ) {

        const step =
            document.getElementById(
                `step${i}`
            );

        step.classList.remove(
            "completed"
        );

        step.classList.remove(
            "active"
        );

    }

}


/* =========================================================
   SIMULATED QR SCAN
   ========================================================= */

function scanQR() {

    const message =
        document.getElementById(
            "verificationMessage"
        );


    message.textContent =
        "✓ QR verified successfully. Now enter OTP.";


    message.style.color =
        "#087e7b";

}


/* =========================================================
   OTP VERIFICATION
   ========================================================= */

function verifyOTP() {

    const enteredOTP =
        document.getElementById(
            "enteredOTP"
        ).value.trim();


    const message =
        document.getElementById(
            "verificationMessage"
        );


    if (!enteredOTP) {

        message.textContent =
            "Please enter OTP.";

        message.style.color =
            "#d93636";

        return;

    }


    if (
        enteredOTP ===
        currentBooking.otp
    ) {

        message.textContent =
            "✓ OTP verified successfully!";

        message.style.color =
            "#087e7b";


        setTimeout(() => {

            deliverLuggage();

        }, 1000);

    }

    else {

        message.textContent =
            "✕ Invalid OTP. Please try again.";

        message.style.color =
            "#d93636";

    }

}


/* =========================================================
   DELIVER LUGGAGE
   ========================================================= */

function deliverLuggage() {

    updateTrackingStep(
        5,
        "DELIVERED"
    );


    document
        .getElementById(
            "deliveredPassenger"
        )
        .textContent =
        currentBooking.passengerName;


    document
        .getElementById(
            "deliveredBooking"
        )
        .textContent =
        currentBooking.bookingId;


    document
        .getElementById(
            "delivered"
        )
        .classList.remove("hidden");


    document
        .getElementById(
            "delivered"
        )
        .scrollIntoView({
            behavior: "smooth"
        });

}


/* =========================================================
   ERROR FUNCTIONS
   ========================================================= */

function showError(message) {

    const error =
        document.getElementById(
            "errorMessage"
        );

    error.textContent =
        message;

}


function clearError() {

    document.getElementById(
        "errorMessage"
    ).textContent = "";

}