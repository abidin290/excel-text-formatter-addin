// Pastikan Office API siap digunakan
Office.onReady(() => {
  console.log("Excel Add-In is ready!");

  // Sambungkan tombol ke fungsi transformasi
  document.getElementById("uppercase").onclick = () => transformText("UPPERCASE");
  document.getElementById("lowercase").onclick = () => transformText("LOWERCASE");
  document.getElementById("titlecase").onclick = () => transformText("TITLECASE");
});

// Fungsi untuk mengubah teks di sel yang dipilih
async function transformText(mode) {
  try {
    console.log("Transforming text to:", mode); // Pindahkan log ke sini

    await Excel.run(async (context) => {
      // Ambil range (sel) yang dipilih
      const range = context.workbook.getSelectedRange();

      // Memuat nilai dari range yang dipilih
      range.load("values");
      await context.sync();

      if (!range.values || range.values.length === 0) {
        console.error("No cells selected.");
        alert("Please select a cell or range of cells before clicking the button.");
        return;
      }

      console.log("Selected range values:", range.values);

      // Transformasi teks berdasarkan mode
      const transformedValues = range.values.map((row) => {
        return row.map((cell) => {
          if (typeof cell === "string") {
            switch (mode) {
              case "UPPERCASE":
                return cell.toUpperCase();
              case "LOWERCASE":
                return cell.toLowerCase();
              case "TITLECASE":
                return cell
                  .toLowerCase()
                  .split(" ")
                  .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
                  .join(" ");
              default:
                return cell;
            }
          }
          return cell; // Biarkan nilai non-string tetap
        });
      });

      // Masukkan kembali nilai yang sudah diubah ke dalam range
      range.values = transformedValues;
      await context.sync();

      console.log(`Text successfully transformed to ${mode}.`);
      alert(`Text successfully transformed to ${mode}.`);
    });
  } catch (error) {
    console.error("Error transforming text:", error);
    alert("An error occurred. Check the console for more details.");
  }
}
