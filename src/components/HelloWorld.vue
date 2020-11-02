<template>
  <div class="hello">
    <button @click="generate()">Сгенерирвоать</button>
  </div>
</template>

<script>
import { Document, Packer, Paragraph, TextRun } from "docx";
import { saveAs } from "file-saver";
export default {
  name: "HelloWorld",
  data() {
    return {};
  },
  computed: {},
  methods: {
    generate() {
      const doc = new Document();

      doc.addSection({
        margins: {
          top: 141.73,
          left: 425.2,
          bottom: 720,
          right: 311.81
        },
        size: {
          width: 3288,
          height: 16837
        },
        properties: {},
        children: [
          ...this.addManufacturer(
            "Carrier – DataCOLD 500",
            "EN12830  T  B  C1, ATP-MUC 1036/1037  TS",
            10
          ),
          this.addCompanyName("VPTRANS", "14", true),
          this.addChassisNumberAndSerialNumber("5078811", "10830059", 14),
          this.addPrintStart("Tuesday 27/08/2020", "22:10:58", 14)
        ]
      });

      Packer.toBlob(doc).then(blob => {
        console.log(blob);
        saveAs(blob, "example.docx");
        console.log("Document created successfully");
      });
    },
    addManufacturer(textFirstLine, textSecondLine, size) {
      return [
        new Paragraph({
          children: [new TextRun({ text: textFirstLine, size: size })]
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: textSecondLine,
              size: size
            })
          ]
        })
      ];
    },
    addCompanyName(text, size, bold) {
      return new Paragraph({
        spacing: {
          before: 140
        },
        children: [
          new TextRun({
            text: text,
            size: size,
            bold: bold
          })
        ]
      });
    },
    addChassisNumberAndSerialNumber(chassisNumber, serialNumber, size) {
      return new Paragraph({
        children: [
          new TextRun({
            text: chassisNumber + ",",
            size: size
          }),
          new TextRun({
            text: "S/N: " + serialNumber,
            size: size
          })
        ]
      });
    },
    addPrintStart(data, time, size) {
      return [
        new Paragraph({
        children: [
          new TextRun({
            text: data,
            size: size
          }),
          new TextRun({
            text: " " + time,
            size: size
          })
        ]
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Sample rate 60minute(s)",
            size: size
          })
        ]
      })
    ];
  }
};
</script>
