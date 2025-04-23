Office.onReady(() => {
  console.log("✅ Office.js ready in functions.js");

  /**
   * Excel custom function: Checks if a ticker is compliant via WordPress API.
   * @customfunction
   * @param {string} ticker The ticker symbol to check.
   * @returns {Promise<string>} The compliance status message.
   */
  async function Compliant(ticker) {
    const apiBase = "https://env-muslimxchange-staging.kinsta.cloud/wp-json/mx/v1";

    try {
      // ✅ Get JWT from shared OfficeRuntime storage
      const token = await OfficeRuntime.storage.getItem("jwt");

      const headers = token
        ? { Authorization: `Bearer ${token}` }
        : {};

      const res = await fetch(`${apiBase}/hello?ticker=${encodeURIComponent(ticker)}`, {
        headers
      });

      if (!res.ok) {
        console.error("❌ API error:", res.status);
        return "❌ API Error";
      }

      const data = await res.json();
      return data.message || `❓ Unknown status for ${ticker}`;
    } catch (err) {
      console.error("❌ Compliance check failed:", err);
      return "❌ Request Failed";
    }
  }

  /**
 * Custom function for TEST
 * @customfunction
 * @param {string} ticker The stock ticker (e.g., "AAPL")
 * @param {...string} fields Fields to return in adjacent columns (e.g., "Result", "AAOIFI")
 * @returns {any[][]} Returns values in a 1-row, N-column array
 */
  function TEST(ticker, ...fields) {
    const mockData = {
      AAPL: {
        Name: "Apple",
        Result: "Q1 +15%",
        AAOIFI: "Compliant"
      },
      TSLA: {
        Name: "Tesla",
        Result: "Q1 -5%",
        AAOIFI: "Non-Compliant"
      }
    };

    const stock = mockData[ticker.toUpperCase()] || { Name: ticker, Result: "N/A", AAOIFI: "N/A" };

    // Flatten any nested field arrays Excel might pass
    const flatFields = fields.flat();

    const result = [stock.Name];
    for (const field of flatFields) {
      result.push(stock[field] || "Unknown");
    }

    return [result];
  }

  CustomFunctions.associate("TEST", TEST);



  CustomFunctions.associate("Compliant", Compliant);
});
