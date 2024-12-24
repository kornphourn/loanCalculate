var app = angular.module('App', []);
app.controller('myCtrl', function ($scope, $sce) {
   $scope.AmountOfMoney = 10000; // Loan Amount
   $scope.DurationOfLoan = 5; // Duration in months (7 years)
   $scope.PercentageOfLoan = 8; // Annual interest rate (%)
   $scope.paymentType = "fixed"; // Default to fixed monthly payments
   $scope.table = "";
   $scope.subTotal = 0;
   $scope.excelData = []; // Data for Excel export

   $scope.rateCalculator = () => {
      // Clear previous calculations
      $scope.table = "";
      $scope.subTotal = 0;
      $scope.excelData = []; 
      // Inputs
      const P = parseFloat($scope.AmountOfMoney); // Principal
      const n = parseInt($scope.DurationOfLoan) * 12; // Total number of payments (months)
      const annualRate = parseFloat($scope.PercentageOfLoan) / 100; // Annual rate in decimal
      const r = annualRate / 12; // Monthly interest rate

      let remainingBalance = P;

      if ($scope.paymentType === "fixed") {
         // Fixed Monthly Payment: Calculate using the amortized formula
         const monthlyPayment = (P * r * Math.pow(1 + r, n)) / (Math.pow(1 + r, n) - 1);

         // Generate the amortization table
         for (let i = 1; i <= n; i++) {
            const interest = remainingBalance * r; // Interest for the current month
            const principal = monthlyPayment - interest; // Principal portion
            $scope.subTotal += monthlyPayment; // Add to total payments

            // Add to table
            $scope.table += `
               <tr>
                  <th scope="row">${i}</th>
                  <td>${remainingBalance.toFixed(2)} $</td>
                  <td>${principal.toFixed(2)} $</td>
                  <td>${interest.toFixed(2)} $</td>
                  <td>${monthlyPayment.toFixed(2)} $</td>
               </tr>
            `;

            // Add to Excel data
            $scope.excelData.push({
               Month: i,
               "Remaining Balance": remainingBalance.toFixed(2),
               "Principal Paid": principal.toFixed(2),
               "Interest Paid": interest.toFixed(2),
               "Total Paid": monthlyPayment.toFixed(2),
            });

            // Update remaining balance
            remainingBalance -= principal;
         }
      } else {
         // Not Fixed Monthly Payment: Calculate interest dynamically
         const accumulate = P / n; // Fixed portion of principal payment each month

         for (let i = 1; i <= n; i++) {
            const interest = remainingBalance * r; // Interest for the current month
            const totalPayment = accumulate + interest; // Total payment for this month
            $scope.subTotal += totalPayment; // Add to total payments

            // Add to table
            $scope.table += `
               <tr>
                  <th scope="row">${i}</th>
                  <td>${remainingBalance.toFixed(2)} $</td>
                  <td>${accumulate.toFixed(2)} $</td>
                  <td>${interest.toFixed(2)} $</td>
                  <td>${totalPayment.toFixed(2)} $</td>
               </tr>
            `;

            // Add to Excel data
            $scope.excelData.push({
               Month: i,
               "Remaining Balance": remainingBalance.toFixed(2),
               "Principal Paid": accumulate.toFixed(2),
               "Interest Paid": interest.toFixed(2),
               "Total Paid": totalPayment.toFixed(2),
            });

            // Update remaining balance
            remainingBalance -= accumulate;
         }
      }

      // Update total payments
      $scope.subTotal = $scope.subTotal.toFixed(2);

      // Render table safely
      $scope.trustedAppTitle = $sce.trustAsHtml($scope.table);
   };

   // Export as Excel
   $scope.exportToExcel = () => {
      // Create a new workbook and worksheet
      const worksheet = XLSX.utils.json_to_sheet($scope.excelData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Loan Schedule");

      // Download the Excel file
      XLSX.writeFile(workbook, "Loan_Schedule.xlsx");
   };
});
