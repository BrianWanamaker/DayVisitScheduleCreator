document.getElementById("visitForm").addEventListener("submit", function (e) {
  e.preventDefault();

  // Gather form data
  const formData = new FormData(e.target);
  const formProps = Object.fromEntries(formData.entries());

  // Create a new document
  const doc = new docx.Document();

  // Add content to the document
  doc.addSection({
    children: [
      new docx.Paragraph({
        text: `Name: ${formProps.studentName}`,
        heading: docx.HeadingLevel.HEADING_1,
      }),
      new docx.Paragraph(`Major: ${formProps.major}`),
      new docx.Paragraph(`Hometown: ${formProps.hometown}`),
      new docx.Paragraph(
        `Admissions Counselor: ${formProps.admissionsCounselor}`
      ),
      new docx.Paragraph(`Host's Name: ${formProps.hostsName}`),
      new docx.Paragraph(`Host's Phone Number: ${formProps.hostsPhoneNumber}`),
      new docx.Paragraph(`Visit Date: ${formProps.visitDate}`),
    ],
  });

  // Generate and download the .docx file
  docx.Packer.toBlob(doc).then((blob) => {
    // Assuming the use of the FileSaver.js library for saving files
    saveAs(blob, `${formProps.studentName}_Visit.docx`);
  });
});
