function serialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);

  const dateStr = new Date(Date.UTC(dateInfo.getFullYear(), dateInfo.getMonth(), dateInfo.getDate()));

  // const serial = 42320;
  // const date = serialToDate(serial);
  // console.log(date);  // Output: YYYY-MM-DD format
  return dateStr.toISOString().split("T")[0];
}

module.exports = {
  serialToDate,
};
