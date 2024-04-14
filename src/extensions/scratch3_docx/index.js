const ArgumentType = require("../../extension-support/argument-type");
const BlockType = require("../../extension-support/block-type");
const Cast = require("../../util/cast");
const log = require("../../util/log");

const blockIconURI =
    "https://png.pngtree.com/png-vector/20190411/ourmid/pngtree-docx-file-document-icon-png-image_927834.jpg";

class Scratch3Docx {
    constructor(runtime) {
        this.runtime = runtime;
        this.docxContent = "";
        this.canvas = document.getElementById("docx");
    }

    getInfo() {
        return {
            id: "docx",
            name: "DocBuilder",
            blockIconURI: blockIconURI,
            blocks: [
                {
                    opcode: "insertHeading",
                    blockType: BlockType.COMMAND,
                    text: "insert heading [TEXT]",
                    arguments: {
                        TEXT: {
                            type: ArgumentType.STRING,
                            defaultValue: "Heading",
                        },
                    },
                    color1: "#3498db", // Set block color to blue
                },
                {
                    opcode: "insertText",
                    blockType: BlockType.COMMAND,
                    text: "insert text [TEXT]",
                    arguments: {
                        TEXT: {
                            type: ArgumentType.STRING,
                            defaultValue: "Text",
                        },
                    },
                    color1: "#3498db", // Set block color to blue
                },
                {
                    opcode: "insertImage",
                    blockType: BlockType.COMMAND,
                    text: "insert image [URL]",
                    arguments: {
                        URL: {
                            type: ArgumentType.STRING,
                            defaultValue: "Image URL",
                        },
                    },
                    color1: "#3498db", // Set block color to blue
                },
                {
                    opcode: "newLine",
                    blockType: BlockType.COMMAND,
                    text: "new line",
                    arguments: {},
                    color1: "#3498db", // Set block color to blue
                },

                {
                    opcode: "make",
                    blockType: BlockType.COMMAND,
                    text: "make [TEXT]",
                    arguments: {
                        TEXT: {
                            type: ArgumentType.DROPDOWN,
                            defaultValue: "B",
                            menu: "boldItalicUnderlineMenu",
                        },
                    },
                    color1: "#3498db", // Set block color to blue
                },
               
                {
                    opcode: "erase",
                    blockType: BlockType.COMMAND,
                    text: "erase",
                    color1: "#3498db", // Set block color to blue
                },
            ],
            menus: {
                boldItalicUnderlineMenu: {
                    items: [
                        { text: "Bold", value: "B" },
                        { text: "Italic", value: "I" },
                        { text: "Underline", value: "U" },
                    ],
                },
                leftRightCenterMenu: {
                    items: [
                        { text: "Left", value: "L" },
                        { text: "Right", value: "R" },
                        { text: "Center", value: "C" },
                    ],
                },
            },
        };
    }

    newLine(args) {
        this.docxContent += "<br/>"; // Add a new line to the document
        this.canvas.innerHTML = this.docxContent;
    }

    insertHeading(args) {
        const text = Cast.toString(args.TEXT);
        this.docxContent += `<h1 style="text-align:center">${text}</h1>`; // Insert a heading with the specified text
        this.canvas.innerHTML = this.docxContent;
    }

    erase() {
        this.docxContent = ""; // Clear the document
        this.canvas.innerHTML = this.docxContent;
    }

    insertText(args) {
        const text = Cast.toString(args.TEXT);
        this.docxContent += `<span>${text}</span>`; // Insert the specified text into the document
        this.canvas.innerHTML = this.docxContent;
    }

    insertImage(args) {
        let imageUrl = args.URL;
        if(imageUrl.startsWith("http")){
            this.docxContent+=`<img src="${imageUrl}" alt="Image" width="100%" height="100%" style="object-fit: cover;" > `
            this.canvas.innerHTML=this.docxContent
            return;
        }
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = 'image/*';
        fileInput.onchange = (event) => {
            const file = event.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = () => {
                    let imageURL = reader.result;
                    this.docxContent+=`<img src="${imageURL}" alt="Image" width="100%" height="100%" style="object-fit: cover;" > `
                    this.canvas.innerHTML=this.docxContent
                    return;
                };
                reader.readAsDataURL(file);
            }
        };
        fileInput.click();
    }

    make(args) {
        const format = Cast.toString(args.TEXT);
        switch (format) {
            case "B":
                this.docxContent += "<strong>";
                this.canvas.innerHTML = this.docxContent;
                break;
            case "I":
                this.docxContent += "<em>";
                this.canvas.innerHTML = this.docxContent;
                break;
            case "U":
                this.docxContent += "<u>";
                this.canvas.innerHTML = this.docxContent;
                break;
        }
    }

   
    saveDocx() {
        // You can implement this method to save the docxContent to a file or perform any other necessary actions
    }
}

module.exports = Scratch3Docx;
