function testGeminiSummary() {
  const result = getAISummary("This is a short test email body.");
  Logger.log("Test summary: " + result);
}
