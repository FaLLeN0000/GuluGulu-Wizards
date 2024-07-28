function calculateMSRP(manufacturingPrice) {
  const minMarkupPercentage = 1.5; // Minimum markup percentage
  const maxMarkupPercentage = 2.0; // Maximum markup percentage
  const markupPercentage = Math.random() * (maxMarkupPercentage - minMarkupPercentage) + minMarkupPercentage;
  return (manufacturingPrice * markupPercentage).toFixed(2);
}

function estimatePromotionPrice(msrp) {
  const discountRate = 0.1; // Example 10% discount
  return (msrp * (1 - discountRate)).toFixed(2);
}

function analyzeSalesAndRecommend(sales, reviews) {
  const threshold = 50; // Example threshold
  return (sales < threshold && reviews < threshold) ? 'Consider Price Change or Halt Import' : 'Good Performance';
}

function createCalendarEvent(productName, msrp, promoPrice, recommendation, eventDate) {
  var calendar = CalendarApp.getDefaultCalendar();
  var eventTitle = `Review: ${productName}`;
  var eventDescription = `Product: ${productName}\nMSRP: ${msrp}\nPromo Price: ${promoPrice}\nRecommendation: ${recommendation}`;

  // Set time to 9 AM
  eventDate.setHours(9);
  eventDate.setMinutes(0);
  eventDate.setSeconds(0);

  calendar.createEvent(eventTitle, eventDate, new Date(eventDate.getTime() + 60 * 60 * 1000), {
    description: eventDescription
  });
}

function processFormResponses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  const data = sheet.getDataRange().getValues();
  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Analysis');

  outputSheet.clearContents();
  outputSheet.appendRow(["Product Name", "MSRP", "Promo Price", "Recommendation"]);

  for (let i = 1; i < data.length; i++) {
    const productType = data[i][7]; // Assuming the 8th column indicates whether the product is new or existing
    const manufacturingPrice = parseFloat(data[i][2]); // Assuming 3rd column is Manufacturing Price
    const sales = parseInt(data[i][3]); // Assuming 4th column is Sales
    const reviews = parseInt(data[i][6]); // Assuming 7th column is Reviews
    const inputDate = new Date(data[i][0]); // Assuming the 1st column is the input date
    const productName = data[i][1]; // Assuming the 2nd column is the product name

    let msrp;
    let promoPrice;
    let recommendation;

    if (productType == "New Product") {
      // For new products, only calculate MSRP and promo price, and set recommendation to "New Product"
      msrp = calculateMSRP(manufacturingPrice);
      promoPrice = "New Product";
      recommendation = "New Product";
    } else if (productType == "Existing Product") {
      // For existing products, proceed with the full analysis
      msrp = calculateMSRP(manufacturingPrice);
      promoPrice = estimatePromotionPrice(msrp);
      recommendation = analyzeSalesAndRecommend(sales, reviews);
    }

    outputSheet.appendRow([productName, msrp, promoPrice, recommendation]);
    createCalendarEvent(productName, msrp, promoPrice, recommendation, inputDate); // Create calendar event with the input date
  }
}
