const ArgumentType = require("../../extension-support/argument-type");
const BlockType = require("../../extension-support/block-type");
const Chart = require("chart.js/auto");

class Scratch3Docx {
    constructor(runtime) {
        this.runtime = runtime;
        this.canvas = document.getElementById("docx");
        this.url = "";
        this.spreadsheetId = "";
        const script = document.createElement("script");
        script.type = "text/javascript";
        script.src = "https://www.gstatic.com/charts/loader.js";
        // Append the script element to the head of the document
        document.head.appendChild(script);
    }
    extractFrom(url) {
        const arr = url.split("/");
        return arr[5];
    }
    convert(columnNumber) {
        let output = "";
        let charCode;
        while (columnNumber >= 1) {
            if (columnNumber % 26 === 0) {
                charCode = 26;
                columnNumber = ~~(columnNumber / 26) - 1;
            } else {
                charCode = columnNumber % 26;
                columnNumber = ~~(columnNumber / 26);
            }
            output = String.fromCharCode(charCode + 64) + output;
        }
        return output;
    }
    async getToken() {
        const code = "Add Your Auth Code using the above url;";
        console.log(code);
        const params = new URLSearchParams({
            code: code,
            client_id: "Client ID",
            client_secret: "Client Secret",
            redirect_uri: "https://127.0.0.1/", // This should match the redirect URI you used in your authorization request
            grant_type: "authorization_code",
        });

        const response = await fetch("https://oauth2.googleapis.com/token", {
            method: "POST",
            headers: {
                "Content-Type": "application/x-www-form-urlencoded",
            },
            body: params.toString(),
        });

        if (response.ok) {
            const data = await response.json();
            const access_token = data.access_token;
            return access_token;
        } else {
            throw new Error("Failed to obtain access token");
        }
    }
    getInfo() {
        return {
            id: "sheets",
            name: "G-Sheets",
            blocks: [
                {
                    opcode: "fetchData",
                    blockType: BlockType.COMMAND,
                    text: "connect to shared sheet [URL]",
                    arguments: {
                        URL: {
                            type: ArgumentType.STRING,
                            defaultValue: "url",
                        },
                    },
                },
                {
                    opcode: "input",
                    blockType: BlockType.COMMAND,
                    text: "input [DATA] to column [COL] and row [ROW]",
                    arguments: {
                        DATA: {
                            type: ArgumentType.NUMBER,
                            defaultValue: 0,
                        },
                        COL: {
                            type: ArgumentType.NUMBER,
                            defaultValue: 0,
                        },
                        ROW: {
                            type: ArgumentType.NUMBER,
                            defaultValue: 0,
                        },
                    },
                },
                {
                    opcode: "read",
                    blockType: BlockType.REPORTER,
                    text: "read cell value at column [COL] and row [ROW]",
                    arguments: {
                        COL: {
                            type: ArgumentType.NUMBER,
                            defaultValue: 0,
                        },
                        ROW: {
                            type: ArgumentType.NUMBER,
                            defaultValue: 0,
                        },
                    },
                },
                {
                    opcode: "plot",
                    blockType: BlockType.COMMAND,
                    text: "Plot [PLOT_TYPE] START[START_ROW] END[END_ROW]",
                    arguments: {
                        PLOT_TYPE: {
                            type: ArgumentType.DROPDOWN,
                            menu: "plotTypeMenu",
                            defaultValue: "Bar",
                        },
                        START_ROW: {
                            type: ArgumentType.STRING,
                            defaultValue: 0,
                        },
                        END_ROW: {
                            type: ArgumentType.STRING,
                            defaultValue: 0,
                        },
                    },
                },
            ],
            menus: {
                plotTypeMenu: {
                    items: [
                        { text: "Bar", value: "Bar" },
                        { text: "Pie", value: "Pie" },
                        { text: "Line", value: "Line" },
                    ],
                },
            },
        };
    }
    async plot(args) {
        const plotType = args.PLOT_TYPE;
        const start = args.START_ROW;
        const end = args.END_ROW;
        const API_KEY = "Add your API Key";

        const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.spreadsheetId}/values/Sheet1!${start}:${end}?key=${API_KEY}`;

        const response = await fetch(url, { method: "GET" });
        const data0 = await response.json();

        console.log(data0);

        const backgroundDiv = document.createElement("div");
        backgroundDiv.style.position = "fixed";
        backgroundDiv.style.top = "0";
        backgroundDiv.style.left = "0";
        backgroundDiv.style.width = "100%";
        backgroundDiv.style.height = "100%";
        backgroundDiv.style.background = "rgba(0, 0, 0, 0.7)"; // Adjust RGBA values for blackish dark background
        backgroundDiv.style.zIndex = "9998"; // Set a lower z-index to be behind fullScreenContainer

        // Append the background div to the body
        backgroundDiv.addEventListener("click", () => {
            document.body.removeChild(fullScreenContainer);
            document.body.removeChild(backgroundDiv);
        });
        document.body.appendChild(backgroundDiv);

        // Generate chart inside a full-screen element
        const fullScreenContainer = document.createElement("div");
        fullScreenContainer.id = "fullScreenContainer";
        fullScreenContainer.style.position = "fixed";
        fullScreenContainer.style.top = "50%";
        fullScreenContainer.style.left = "50%";
        fullScreenContainer.style.transform = "translate(-50%, -50%)"; // Center the container
        fullScreenContainer.style.width = "50%";
        fullScreenContainer.style.height = "50%";
        fullScreenContainer.style.background = "white"; // Set background to white for the
        fullScreenContainer.style.border = "2px solid black"; // Add border
        fullScreenContainer.style.zIndex = "9999";
        fullScreenContainer.style.display = "flex";
        fullScreenContainer.style.alignItems = "center";
        fullScreenContainer.style.justifyContent = "center";

        const canvas = document.createElement("canvas");
        canvas.id = "myChart";
        canvas.width = window.innerWidth; // Set canvas width to window width
        canvas.height = window.innerHeight; // Set canvas height to window height

        fullScreenContainer.appendChild(canvas);

        // Add close button
        const closeSymbol = document.createElement("div");
        closeSymbol.textContent = "✖"; // Use the "×" character for a cross symbol
        closeSymbol.style.position = "absolute";
        closeSymbol.style.top = "1px";
        closeSymbol.style.right = "10px";
        closeSymbol.style.zIndex = "9999";
        closeSymbol.style.fontSize = "24px";
        closeSymbol.style.cursor = "pointer";
        closeSymbol.addEventListener("click", () => {
            document.body.removeChild(fullScreenContainer);
            document.body.removeChild(backgroundDiv);
        });

        fullScreenContainer.appendChild(closeSymbol);

        document.body.appendChild(fullScreenContainer);

        // Draw chart
        var ctx = canvas.getContext("2d");
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        var chartType = plotType.toLowerCase(); // Convert the plotType to lowercase for consistency

        // Sample data (replace with your actual data)
        var range = response.result;
        var spreadsheetLabel = [];
        var spreadsheetData = [];
        if (data0.values.length > 0) {
            // Process the fetched data
            data0.values.forEach(function (row) {
                spreadsheetLabel.push(row[0]);
                spreadsheetData.push(parseFloat(row[1])); // Assuming the second column contains numerical data
            });
        } else {
            console.log("No data found.");
        }
        console.log(spreadsheetLabel, spreadsheetData);
        var data = {
            labels: spreadsheetLabel,
            datasets: [
                {
                    label: "My Dataset",
                    data: spreadsheetData,
                    backgroundColor: [
                        "rgba(255, 99, 132, 0.2)",
                        "rgba(54, 162, 235, 0.2)",
                        "rgba(255, 206, 86, 0.2)",
                        "rgba(75, 192, 192, 0.2)",
                        "rgba(153, 102, 255, 0.2)",
                        "rgba(255, 159, 64, 0.2)",
                    ],
                    borderColor: [
                        "rgba(255, 99, 132, 1)",
                        "rgba(54, 162, 235, 1)",
                        "rgba(255, 206, 86, 1)",
                        "rgba(75, 192, 192, 1)",
                        "rgba(153, 102, 255, 1)",
                        "rgba(255, 159, 64, 1)",
                    ],
                    borderWidth: 1,
                },
            ],
        };

        // Configuration options
        var options = {
            scales: {
                y: {
                    beginAtZero: true,
                },
            },
        };
        // Choose chart type based on selected plot type
        var chartTypeMap = {
            bar: "bar",
            pie: "pie",
            line: "line",
        };
        if (this.myChart) {
            this.myChart.destroy();
        }

        this.myChart = new Chart(ctx, {
            type: chartTypeMap[chartType],
            data: data,
            options: options,
        });
    }

    async fetchData(args) {
        this.url = args.URL;
        const url = this.url;
        console.log(url);
        fetch(url, {
            method: "GET",
        })
            .then(() => {
                this.spreadsheetId = this.extractFrom(this.url);
                alert("Connected SuccessFully ");
            })
            .catch((err) => {
                alert("Wrong URL!!");
                console.log(err);
            });
    }
    async input(args) {
        const col = args.COL;
        const row = args.ROW;
        const value = args.DATA;
        const convertedCol = this.convert(col);
        // const API_KEY = 'YOUR_API_KEY'; // This is not needed for authorized requests
        const ACCESS_TOKEN = await this.getToken();

        const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.spreadsheetId}/values/Sheet1!${convertedCol}${row}:${convertedCol}${row}?valueInputOption=RAW`;

        const response = await fetch(url, {
            method: "PUT",
            headers: {
                Authorization: `Bearer ${ACCESS_TOKEN}`,
                "Content-Type": "application/json",
            },
            body: JSON.stringify({
                range: `Sheet1!${convertedCol}${row}:${convertedCol}${row}`,
                majorDimension: "ROWS",
                values: [[value]],
            }),
        });

        const data = await response.json();
        console.log(data);
        if (!data) {
            this.canvas.innerHTML = `<h2>Wrong Input</h2>`;
            return;
        }
        this.canvas.innerHTML = `<p>${data}<p/>`;
    }
    async read(args) {
        const col = args.COL;
        const row = args.ROW;
        const convertedCol = this.convert(col);
        const API_KEY = "Add your API key";
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.spreadsheetId}/values/Sheet1!${convertedCol}${row}:${convertedCol}${row}?key=${API_KEY}`;
        const response = await fetch(url, { method: "GET" });
        const data = await response.json();
        if (!data.values) {
            this.canvas.innerHTML = `<h2>Wrong Input</h2>`;
            return;
        }
        this.canvas.innerHTML = `<p>The value is :<p/><h2>${data.values[0][0]}</h2>`;
    }
}

module.exports = Scratch3Docx;
