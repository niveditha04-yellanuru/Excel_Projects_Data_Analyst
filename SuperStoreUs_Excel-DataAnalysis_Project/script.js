/* =========================================
   GLOBAL VARIABLES
========================================= */

let rawData = [];

let filteredData = [];

let charts = {};

if(typeof ChartDataLabels !== "undefined") {
    Chart.register(ChartDataLabels);
}

Chart.defaults.layout.padding = 0;
Chart.defaults.plugins.legend.labels.padding = 4;
Chart.defaults.plugins.tooltip.enabled = true;
Chart.defaults.plugins.tooltip.events = ["click"];
Chart.defaults.events = ["click"];
Chart.defaults.interaction.mode = "index";
Chart.defaults.interaction.intersect = false;
Chart.defaults.plugins.datalabels.display = false;
Chart.defaults.scales.linear.ticks.display = false;
Chart.defaults.animation = false;
Chart.defaults.transitions.active.animation.duration = 0;
Chart.defaults.elements.arc.hoverOffset = 0;
Chart.defaults.elements.arc.hoverBorderWidth = 0;


/* =========================================
   DEFAULT DEMO DATA

   Website works before uploading dataset.
========================================= */

rawData = [

{
"Order Priority":"High",
"Customer Segment":"Consumer",
"Sales":37000,
"Profit":6000,
"Discount":0.05,
"Product Sub-Category":"Binders",
"Ship Mode":"Regular Air",
"Shipping Cost":800,
"Order Date":"2024-01-10",
"State":"California",
"Order ID":"SS001"
},

{
"Order Priority":"Low",
"Customer Segment":"Corporate",
"Sales":38000,
"Profit":5000,
"Discount":0.04,
"Product Sub-Category":"Chairs & Chairmats",
"Ship Mode":"Delivery Truck",
"Shipping Cost":950,
"Order Date":"2024-02-12",
"State":"Texas",
"Order ID":"SS002"
},

{
"Order Priority":"Medium",
"Customer Segment":"Consumer",
"Sales":29000,
"Profit":7200,
"Discount":0.06,
"Product Sub-Category":"Labels",
"Ship Mode":"Express Air",
"Shipping Cost":600,
"Order Date":"2024-03-15",
"State":"New York",
"Order ID":"SS003"
},

{
"Order Priority":"Critical",
"Customer Segment":"Home Office",
"Sales":45000,
"Profit":12500,
"Discount":0.03,
"Product Sub-Category":"Office Machines",
"Ship Mode":"Regular Air",
"Shipping Cost":1200,
"Order Date":"2024-04-20",
"State":"Florida",
"Order ID":"SS004"
},

{
"Order Priority":"High",
"Customer Segment":"Corporate",
"Sales":32000,
"Profit":8500,
"Discount":0.05,
"Product Sub-Category":"Telephones",
"Ship Mode":"Express Air",
"Shipping Cost":700,
"Order Date":"2024-05-05",
"State":"Washington",
"Order ID":"SS005"
},

{
"Order Priority":"Critical",
"Customer Segment":"Consumer",
"Sales":51000,
"Profit":14000,
"Discount":0.08,
"Product Sub-Category":"Binders",
"Ship Mode":"Delivery Truck",
"Shipping Cost":1400,
"Order Date":"2024-06-10",
"State":"Illinois",
"Order ID":"SS006"
}

];

filteredData = [...rawData];


/* =========================================
   COLUMN DETECTION

   Makes dashboard compatible with
   different SuperStore datasets.
========================================= */

function findColumn(possibleNames) {

    if (!filteredData.length) return null;

    let columns =
    Object.keys(filteredData[0]);

    for (let name of possibleNames) {

        let found =
        columns.find(col =>

            col.toLowerCase()
            .replace(/[^a-z]/g,"")
            .includes(

                name.toLowerCase()
                .replace(/[^a-z]/g,"")

            )

        );

        if(found) return found;

    }

    return null;

}


function getNumber(value) {

    if(value === null || value === undefined)
        return 0;

    if(typeof value === "number")
        return value;

    return Number(

        String(value)
        .replace(/[$,%]/g,"")
        .replace(/,/g,"")

    ) || 0;

}


function cleanLabel(value) {

    return String(value ?? "Unknown")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

}


function getMonthLabel(value) {

    if(value === null || value === undefined || value === "")
        return null;

    if(typeof value === "number") {

        let date =
        new Date(Date.UTC(1899, 11, 30) + value * 86400000);

        return date.toLocaleString("en-US", {
            month:"short",
            timeZone:"UTC"
        });

    }

    let text =
    String(value).trim();

    let parsed =
    new Date(text);

    if(!Number.isNaN(parsed.getTime())) {

        return parsed.toLocaleString("en-US", {
            month:"short",
            timeZone:"UTC"
        });

    }

    let month =
    text.match(/\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*/i);

    return month ? month[1].slice(0, 1).toUpperCase() + month[1].slice(1, 3).toLowerCase() : null;

}


/* =========================================
   LOAD CSV / EXCEL
========================================= */

function loadFile(event) {

    let file = event.target.files[0];

    if(!file) return;


    showToast(
        "Reading dataset..."
    );


    let reader =
    new FileReader();


    reader.onload =
    function(e) {

        let workbook =
        XLSX.read(

            e.target.result,

            {
                type:"array"
            }

        );


        let firstSheet =
        workbook.SheetNames[0];


        let worksheet =
        workbook.Sheets[firstSheet];


        rawData =
        XLSX.utils.sheet_to_json(
            worksheet,
            {
                defval:""
            }
        );


        filteredData =
        [...rawData];


        buildFilters();

        updateDashboard();


        showToast(
            rawData.length +
            " records loaded successfully!"
        );

    };


    reader.readAsArrayBuffer(file);

}


/* =========================================
   BUILD SIDEBAR FILTERS
========================================= */

function buildFilters() {

    let priorityCol =
    findColumn([
        "Order Priority",
        "Priority"
    ]);


    let segmentCol =
    findColumn([
        "Customer Segment",
        "Segment"
    ]);


    if(priorityCol) {

        let values =
        [...new Set(
            rawData.map(
                x => x[priorityCol]
            )
        )];


        document.getElementById(
            "priorityFilters"
        ).innerHTML =

        values.map(value => `

            <label class="check-item">

            <input
            type="checkbox"
            class="priorityCheck"
            value="${value}"
            checked
            onchange="updateDashboard()">

            ${value}

            </label>

        `).join("");

    }


    if(segmentCol) {

        let values =
        [...new Set(
            rawData.map(
                x => x[segmentCol]
            )
        )];


        document.getElementById(
            "segmentFilters"
        ).innerHTML =

        values.map(value => `

            <label class="check-item">

            <input
            type="checkbox"
            class="segmentCheck"
            value="${value}"
            checked
            onchange="updateDashboard()">

            ${value}

            </label>

        `).join("");

    }

}


/* =========================================
   APPLY FILTERS
========================================= */

function applyFilters() {

    let priorityCol =
    findColumn([
        "Order Priority",
        "Priority"
    ]);


    let segmentCol =
    findColumn([
        "Customer Segment",
        "Segment"
    ]);


    let profitCol =
    findColumn([
        "Profit"
    ]);


    let selectedPriority =
    [...document.querySelectorAll(
        ".priorityCheck:checked"
    )].map(x => x.value);


    let selectedSegment =
    [...document.querySelectorAll(
        ".segmentCheck:checked"
    )].map(x => x.value);


    let positiveOnly =
    document.getElementById(
        "positiveProfit"
    ).checked;


    filteredData =
    rawData.filter(row => {

        let priorityOK =
        !priorityCol ||
        selectedPriority.length === 0 ||
        selectedPriority.includes(
            String(row[priorityCol])
        );


        let segmentOK =
        !segmentCol ||
        selectedSegment.length === 0 ||
        selectedSegment.includes(
            String(row[segmentCol])
        );


        let profitOK =
        !positiveOnly ||
        !profitCol ||
        getNumber(row[profitCol]) > 0;


        return (
            priorityOK &&
            segmentOK &&
            profitOK
        );

    });

}


/* =========================================
   MAIN UPDATE
========================================= */

function updateDashboard() {

    applyFilters();

    updateKPIs();

    createPriorityChart();

    createProfitChart();

    createSubcategoryChart();

    createSegmentChart();

    createShippingChart();

    createMap();

    generateInsights();

    renderTable();

}


/* =========================================
   KPI CALCULATIONS
========================================= */

function updateKPIs() {

    let salesCol =
    findColumn(["Sales"]);

    let profitCol =
    findColumn(["Profit"]);

    let discountCol =
    findColumn(["Discount"]);

    let orderCol =
    findColumn([
        "Order ID",
        "OrderID"
    ]);


    let sales =
    filteredData.reduce(
        (sum,row) =>
        sum + getNumber(row[salesCol]),
        0
    );


    let profit =
    filteredData.reduce(
        (sum,row) =>
        sum + getNumber(row[profitCol]),
        0
    );


    let discount =
    filteredData.length
    ?
    filteredData.reduce(
        (sum,row) =>
        sum + getNumber(row[discountCol]),
        0
    )
    /
    filteredData.length
    :
    0;


    let orders =
    orderCol

    ?

    new Set(
        filteredData.map(
            x => x[orderCol]
        )
    ).size

    :

    filteredData.length;


    let margin =
    sales
    ?
    (profit / sales) * 100
    :
    0;


    let avgOrder =
    orders
    ?
    sales / orders
    :
    0;


    document.getElementById("totalSales").innerText =
    formatKpiCurrency(sales);

    document.getElementById("totalProfit").innerText =
    formatKpiCurrency(profit);


    document.getElementById(
        "avgDiscount"
    ).innerText =
    (discount * 100).toFixed(2) + "%";


    document.getElementById(
        "profitMargin"
    ).innerText =
    margin.toFixed(1) + "%";


    document.getElementById(
        "avgOrder"
    ).innerText =
    Math.round(avgOrder).toLocaleString(
        undefined,
        {
            maximumFractionDigits:0
        }
    );

}


function formatKpiCurrency(value) {

    let rounded =
    Math.round(value);

    if(rounded >= 1000000) {
        return "$" + (rounded / 1000000).toFixed(2) + "M";
    }

    return "$" + rounded.toLocaleString();

}


/* =========================================
   KPI ANIMATION
========================================= */

function animateValue(id,value,prefix="") {

    let element =
    document.getElementById(id);


    let start = 0;

    let duration = 600;

    let startTime = null;


    function animation(time) {

        if(!startTime)
            startTime = time;


        let progress =
        Math.min(
            (time-startTime)/duration,
            1
        );


        let current =
        start +
        (value-start)*progress;


        element.innerText =
        prefix +
        Math.round(current)
        .toLocaleString();


        if(progress < 1) {

            requestAnimationFrame(
                animation
            );

        }

    }


    requestAnimationFrame(animation);

}


/* =========================================
   DESTROY CHART
========================================= */

function destroyChart(name) {

    if(charts[name]) {

        charts[name].destroy();

    }

}


/* =========================================
   PRIORITY CHART
========================================= */

function createPriorityChart() {

    let priorityCol =
    findColumn([
        "Order Priority",
        "Priority"
    ]);

    let salesCol =
    findColumn(["Sales"]);

    let profitCol =
    findColumn(["Profit"]);


    if(!priorityCol) return;


    let groups = {};


    filteredData.forEach(row => {

        let key =
        cleanLabel(row[priorityCol]);


        if(!groups[key]) {

            groups[key] =
            {
                sales:0,
                profit:0
            };

        }


        groups[key].sales +=
        getNumber(row[salesCol]);


        groups[key].profit +=
        getNumber(row[profitCol]);

    });


    destroyChart("priority");


    charts.priority =
    new Chart(

        document.getElementById(
            "priorityChart"
        ),

        {

            type:"bar",

            data:{

                labels:
                Object.keys(groups),

                datasets:[

                {

                    label:"Sales",

                    data:
                    Object.values(groups)
                    .map(x => x.sales),

                    backgroundColor:
                    "#4f79a8",

                    borderRadius:6

                },

                {

                    label:"Profit",

                    type:"line",

                    data:
                    Object.values(groups)
                    .map(x => x.profit),

                    borderColor:
                    "#c44e4b",

                    backgroundColor:
                    "#93ad4d",

                    pointBackgroundColor:
                    "#93ad4d",

                    pointBorderColor:
                    "#93ad4d",

                    pointRadius:3,

                    borderWidth:2,

                    tension:0.25,

                    fill:false

                }

                ]

            },

            options:{

                responsive:true,

                maintainAspectRatio:false,

                layout:{
                    padding:0
                },

                datasets:{
                    bar:{
                        categoryPercentage:.8,
                        barPercentage:.85
                    }
                },

                scales:{
                    y:{
                        beginAtZero:true
                    }
                },

                plugins:{

                    legend:{

                        labels:{

                            boxWidth:10

                        }

                    }

                }

            }

        }

    );

}


/* =========================================
   PROFIT PIE CHART
========================================= */

function createProfitChart() {

    let discountCol =
    findColumn([
        "Discount"
    ]);

    let profitCol =
    findColumn(["Profit"]);


    if(!discountCol) return;


    let groups = {};


    filteredData.forEach(row => {

        let discount =
        getNumber(row[discountCol]);

        let key =
        String(discount);


        groups[key] =
        (groups[key] || 0)
        +
        getNumber(row[profitCol]);

    });


    destroyChart("profit");


    charts.profit =
    new Chart(

        document.getElementById(
            "profitChart"
        ),

        {

            type:"pie",

            data:{

                labels:
                Object.keys(groups),

                datasets:[{

                    data:
                    Object.values(groups),

                    backgroundColor:[

                        "#4f79a8",
                        "#c44e4b",
                        "#93ad4d",
                        "#e1a54b",
                        "#7b61a8"

                    ],

                    borderWidth:0

                }]

            },

            options:{

                responsive:true,

                maintainAspectRatio:false,

                plugins:{

                    datalabels:{
                        display:true,
                        color:"#1f2937",
                        formatter:value => "$" + value.toLocaleString("en-US", {
                            minimumFractionDigits:2,
                            maximumFractionDigits:2
                        }),
                        anchor:"end",
                        align:"end",
                        offset:4,
                        clamp:true
                    },

                    legend:{

                        display:false,

                        position:"right",

                        labels:{

                            boxWidth:10

                        }

                    }

                }

            }

        }

    );

}


/* =========================================
   TOP SUBCATEGORY
========================================= */

function createSubcategoryChart() {

    let categoryCol =
    findColumn([
        "Product Sub-Category",
        "Sub-Category",
        "Sub Category"
    ]);

    let profitCol =
    findColumn(["Profit"]);


    if(!categoryCol) return;


    let groups = {};


    filteredData.forEach(row => {

        let key =
        cleanLabel(row[categoryCol]);


        groups[key] =
        (groups[key] || 0)
        +
        getNumber(row[profitCol]);

    });


    let top =
    Object.entries(groups)

    .sort(
        (a,b) => b[1]-a[1]
    )

    .slice(0,5);


    destroyChart("subcategory");


    charts.subcategory =
    new Chart(

        document.getElementById(
            "subcategoryChart"
        ),

        {

            type:"bar",

            data:{

                labels:
                top.map(x => x[0]),

                datasets:[{

                    label:"Profit",

                    data:
                    top.map(x => x[1]),

                    backgroundColor:
                    "#4f79a8",

                    borderRadius:6

                }]

            },

            options:{

                indexAxis:"y",

                responsive:true,

                maintainAspectRatio:false,

                plugins:{

                    legend:{
                        display:false
                    }

                }

            }

        }

    );

}


/* =========================================
   CUSTOMER SEGMENT
========================================= */

function createSegmentChart() {

    let segmentCol =
    findColumn([
        "Customer Segment",
        "Segment"
    ]);

    let salesCol =
    findColumn(["Sales"]);

    let profitCol =
    findColumn(["Profit"]);


    if(!segmentCol) return;


    let groups = {};


    filteredData.forEach(row => {

        let key =
        row[segmentCol] || "Unknown";


        if(!groups[key]) {

            groups[key] =
            {
                sales:0,
                profit:0
            };

        }


        groups[key].sales +=
        getNumber(row[salesCol]);


        groups[key].profit +=
        getNumber(row[profitCol]);

    });


    destroyChart("segment");


    charts.segment =
    new Chart(

        document.getElementById(
            "segmentChart"
        ),

        {

            type:"line",

            data:{

                labels:
                Object.keys(groups),

                datasets:[

                {

                    label:"Sales",

                    data:
                    Object.values(groups)
                    .map(x => x.sales),

                    borderColor:
                    "#4f79a8",

                    backgroundColor:
                    "rgba(79,121,168,.15)",

                    fill:true,

                    tension:.4

                },

                {

                    label:"Profit",

                    data:
                    Object.values(groups)
                    .map(x => x.profit),

                    borderColor:
                    "#c44e4b",

                    tension:.4

                }

                ]

            },

            options:{

                responsive:true,

                maintainAspectRatio:false

            }

        }

    );

}


/* =========================================
   SHIPPING CHART
========================================= */

function createShippingChart() {

    let shipCol =
    findColumn([
        "Ship Mode"
    ]);

    let costCol =
    findColumn([
        "Shipping Cost",
        "Ship Cost"
    ]);

    let dateCol =
    findColumn([
        "Order Date",
        "Ship Date",
        "Date"
    ]);

    let monthCol =
    findColumn([
        "Month"
    ]);


    if(!shipCol || !costCol) {

        document.getElementById(
            "shippingChart"
        ).style.opacity = .4;

        return;

    }


    let groups = {};
    let modes = [];
    let monthOrder = {};
    let monthNames = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ];


    filteredData.forEach(row => {

        let mode =
        cleanLabel(row[shipCol]);

        let month =
        getMonthLabel(dateCol ? row[dateCol] : row[monthCol]);

        if(!month) return;

        if(!groups[month]) {
            groups[month] = {};
        }

        if(!modes.includes(mode)) {
            modes.push(mode);
        }

        monthOrder[month] = monthNames.indexOf(month);

        groups[month][mode] =
        (groups[month][mode] || 0) + getNumber(row[costCol]);

    });

    let labels =
    Object.keys(groups).sort((a,b) => monthOrder[a] - monthOrder[b]);


    destroyChart("shipping");


    charts.shipping =
    new Chart(

        document.getElementById(
            "shippingChart"
        ),

        {

            type:"line",

            data:{

                labels,

                datasets:modes.map((mode, index) => ({

                    label:mode,

                    data:labels.map(month => groups[month][mode] || 0),

                    borderColor:["#4f79a8", "#c44e4b", "#93ad4d", "#e1a54b", "#7b61a8"][index % 5],

                    backgroundColor:["rgba(79,121,168,.9)", "rgba(196,78,75,.9)", "rgba(147,173,77,.9)", "rgba(225,165,75,.9)", "rgba(123,97,168,.9)"][index % 5],

                    fill:true,

                    tension:.25,

                    pointRadius:0,

                    borderWidth:1,

                    stack:"shipping"

                }))

            },

            options:{

                responsive:true,

                maintainAspectRatio:false,

                interaction:{
                    mode:"index",
                    intersect:false
                },

                scales:{

                    x:{
                        stacked:true
                    },

                    y:{
                        stacked:true,
                        beginAtZero:true,
                        ticks:{
                            callback:value => "$" + Number(value).toLocaleString()
                        }
                    }

                },

                plugins:{

                    legend:{
                        display:true,
                        position:"bottom",
                        labels:{
                            boxWidth:10
                        }
                    }

                }

            }

        }

    );

}


/* =========================================
   USA STATE MAP
========================================= */

const stateCodes = {

"Alabama":"AL",
"Alaska":"AK",
"Arizona":"AZ",
"Arkansas":"AR",
"California":"CA",
"Colorado":"CO",
"Connecticut":"CT",
"Delaware":"DE",
"Florida":"FL",
"Georgia":"GA",
"Hawaii":"HI",
"Idaho":"ID",
"Illinois":"IL",
"Indiana":"IN",
"Iowa":"IA",
"Kansas":"KS",
"Kentucky":"KY",
"Louisiana":"LA",
"Maine":"ME",
"Maryland":"MD",
"Massachusetts":"MA",
"Michigan":"MI",
"Minnesota":"MN",
"Mississippi":"MS",
"Missouri":"MO",
"Montana":"MT",
"Nebraska":"NE",
"Nevada":"NV",
"New Hampshire":"NH",
"New Jersey":"NJ",
"New Mexico":"NM",
"New York":"NY",
"North Carolina":"NC",
"North Dakota":"ND",
"Ohio":"OH",
"Oklahoma":"OK",
"Oregon":"OR",
"Pennsylvania":"PA",
"Rhode Island":"RI",
"South Carolina":"SC",
"South Dakota":"SD",
"Tennessee":"TN",
"Texas":"TX",
"Utah":"UT",
"Vermont":"VT",
"Virginia":"VA",
"Washington":"WA",
"West Virginia":"WV",
"Wisconsin":"WI",
"Wyoming":"WY"

};


function createMap() {

    let stateCol =
    findColumn([
        "State",
        "State or Province"
    ]);

    let salesCol =
    findColumn(["Sales"]);


    if(!stateCol || !salesCol) return;


    let groups = {};


    filteredData.forEach(row => {

        let state =
        row[stateCol];


        if(!state) return;


        groups[state] =
        (groups[state] || 0)
        +
        getNumber(row[salesCol]);

    });


    let locations =
    Object.keys(groups)

    .map(state =>
        stateCodes[state]
    )

    .filter(Boolean);


    let values =
    Object.keys(groups)

    .filter(state =>
        stateCodes[state]
    )

    .map(state =>
        groups[state]
    );


    Plotly.newPlot(

        "salesMap",

        [{

            type:"choropleth",

            locationmode:
            "USA-states",

            locations:

            locations,

            z:

            values,

            text:

            Object.keys(groups),

            colorscale:
            "Blues",

            colorbar:{

                title:"Sales"

            }

        }],

        {

            geo:{

                scope:"usa",

                bgcolor:
                "rgba(0,0,0,0)",

                lakecolor:
                "rgba(0,0,0,0)"

            },

            margin:{

                t:0,
                b:0,
                l:0,
                r:0

            },

            paper_bgcolor:
            "rgba(0,0,0,0)",

            plot_bgcolor:
            "rgba(0,0,0,0)"

        },

        {

            responsive:true,

            displayModeBar:false

        }

    );

}


/* =========================================
   SMART BUSINESS INSIGHTS
========================================= */

function generateInsights() {

    let salesCol =
    findColumn(["Sales"]);

    let profitCol =
    findColumn(["Profit"]);

    let categoryCol =
    findColumn([
        "Product Sub-Category",
        "Sub-Category"
    ]);


    let totalSales =
    filteredData.reduce(
        (sum,row) =>
        sum + getNumber(row[salesCol]),
        0
    );


    let totalProfit =
    filteredData.reduce(
        (sum,row) =>
        sum + getNumber(row[profitCol]),
        0
    );


    let bestCategory = "";

    let groups = {};


    if(categoryCol) {

        filteredData.forEach(row => {

            let category =
            row[categoryCol];


            groups[category] =
            (groups[category] || 0)
            +
            getNumber(row[profitCol]);

        });


        bestCategory =
        Object.entries(groups)

        .sort(
            (a,b)=>b[1]-a[1]
        )[0];

    }


    document.getElementById(
        "insights"
    ).innerHTML = `

        <div class="insight-item">

        <div>

        <strong>
        Revenue Performance
        </strong>

        <br>

        Current filtered data generated
        <b>$${Math.round(totalSales).toLocaleString()}</b>
        in total sales.

        </div>

        </div>


        <div class="insight-item">

        <div>

        <strong>
        Profit Analysis
        </strong>

        <br>

        The business generated
        <b>$${Math.round(totalProfit).toLocaleString()}</b>
        in profit.

        </div>

        </div>


        <div class="insight-item">

        <div>

        <strong>
        Top Opportunity
        </strong>

        <br>

        ${
            bestCategory

            ?

            `<b>${bestCategory[0]}</b>
            is currently the strongest
            profit-performing sub-category.`

            :

            "Upload more detailed data to identify top categories."

        }

        </div>

        </div>


        <div class="insight-item">

        <div>

        <strong>
        Analyst Recommendation
        </strong>

        <br>

        Focus on high-profit categories,
        optimize low-performing segments,
        and monitor discount impact on profitability.

        </div>

        </div>

    `;

}


/* =========================================
   DATA TABLE
========================================= */

function renderTable() {

    if(!filteredData.length) return;


    let search =
    document.getElementById(
        "searchBox"
    ).value.toLowerCase();


    let data =
    filteredData.filter(row =>

        Object.values(row)

        .join(" ")

        .toLowerCase()

        .includes(search)

    );


    let columns =
    Object.keys(filteredData[0]);


    document.getElementById(
        "tableHead"
    ).innerHTML =

    "<tr>" +

    columns.map(col =>
        `<th>${col}</th>`
    ).join("")

    +

    "</tr>";


    document.getElementById(
        "tableBody"
    ).innerHTML =

    data

    .slice(0,100)

    .map(row => `

        <tr>

        ${columns.map(col =>

            `<td>${row[col]}</td>`

        ).join("")}

        </tr>

    `)

    .join("");

}


/* =========================================
   RESET FILTERS
========================================= */

function clearFilters() {

    document

    .querySelectorAll(
        ".priorityCheck,.segmentCheck"
    )

    .forEach(x => x.checked = true);


    document.getElementById(
        "positiveProfit"
    ).checked = false;


    updateDashboard();

    showToast(
        "Filters reset successfully!"
    );

}


/* =========================================
   REFRESH
========================================= */

function refreshDashboard() {

    let button =
    document.querySelector(
        ".refresh"
    );


    button.innerText =
    "Updating...";


    button.disabled = true;


    setTimeout(() => {

        updateDashboard();

        button.innerText =
            "Refresh";

        button.disabled = false;

        showToast(
            "Dashboard updated!"
        );

    },700);

}


/* =========================================
   DOWNLOAD FILTERED CSV
========================================= */

function downloadCSV() {

    if(!filteredData.length) return;


    let worksheet =
    XLSX.utils.json_to_sheet(
        filteredData
    );


    let csv =
    XLSX.utils.sheet_to_csv(
        worksheet
    );


    let blob =
    new Blob(
        [csv],
        {
            type:"text/csv"
        }
    );


    let url =
    URL.createObjectURL(blob);


    let a =
    document.createElement("a");


    a.href = url;

    a.download =
    "SuperStore_Filtered_Data.csv";


    a.click();


    showToast(
        "Filtered data exported!"
    );

}


/* =========================================
   DARK MODE
========================================= */

function toggleTheme() {

    document.body.classList.toggle(
        "dark"
    );


    showToast(
        document.body.classList.contains("dark")

        ?

        "Dark mode enabled"

        :

        "Light mode enabled"

    );

}


/* =========================================
   CLOCK
========================================= */

let liveStartTime = Date.now();


function updateClock() {

    let elapsedSeconds =
    Math.floor((Date.now() - liveStartTime) / 1000);

    let hours =
    Math.floor(elapsedSeconds / 3600)
    .toString()
    .padStart(2, "0");

    let minutes =
    Math.floor((elapsedSeconds % 3600) / 60)
    .toString()
    .padStart(2, "0");

    let seconds =
    (elapsedSeconds % 60)
    .toString()
    .padStart(2, "0");

    document.getElementById("clock").innerText =
    `${hours}:${minutes}:${seconds}`;

}


setInterval(
    updateClock,
    1000
);


updateClock();


/* =========================================
   TOAST
========================================= */

function showToast(message) {

    let toast =
    document.getElementById(
        "toast"
    );


    toast.innerText =
    message;


    toast.classList.add(
        "show"
    );


    setTimeout(() => {

        toast.classList.remove(
            "show"
        );

    },3000);

}


/* =========================================
   INITIALIZE
========================================= */

window.onload = function() {

    setTimeout(() => {

        document

        .getElementById(
            "loader"
        )

        .classList.add(
            "hidden"
        );

    },1200);


    buildFilters();

    updateDashboard();

};
