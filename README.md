# Assignment-2-Data-Cleaning-Transformation
First, check the dataset for missing values in the **Price** column and decide how to handle products with missing prices. If any **Category** values are missing, an appropriate strategy should be used to fill or deal with them. Next, identify inconsistent text formats in the **Product Name** column and any spelling mistakes in the **Category** column, then use **Find and Replace** to standardize the product names and correct category errors. After that, identify and remove any duplicate rows based on the complete row data. Split the **Product ID** column into separate **Manufacturing Date** and **Country Code** columns, removing any unnecessary characters, and merge the **Brand Name** and **Product Name** columns into a new **Product Brand** column. Format the **Price** column as currency and display the **Manufacturing Date** in **DD-MM-YYYY** format. Finally, apply data bars or color scales to the **Price** column and create a custom conditional formatting rule to highlight cells where the **Category** is **Electronics**

### 1. Handling Missing Values & Correcting Inconsistent Data

* Used **Power Query – Clean, Trim, and Capitalize Each Word** to clean and standardize the dataset.
* Used **Statistics → Median** to find the median price, which was **130**.
* Filled the **3 missing Price values** with the median value of **130**.
* For missing **Category** values, used the **Product Name and related information** to determine the appropriate category.
* Identified capitalization inconsistencies in the **Product Name** column.
* Used **Find and Replace** to standardize the Product Names.
* Identified and corrected **typos in the Category** column.

### 2. Removing Duplicates

* Checked the dataset for duplicate rows.
* Used Excel's **Remove Duplicates** feature to identify and remove duplicate rows.
* Removed duplicates based on the **complete row data**.

### 3. Splitting and Merging Data

* Modified the **Product ID** to include the year as instructed.
* Used Excel's **LEFT()** and **RIGHT()** functions to extract:

  * **Manufacturing Date:** `=LEFT([@[Product ID]],11)`
  * **Country Code:** `=RIGHT([@[Product ID]],2)`
* Initially tried to use **Text to Columns**, but some errors occurred, so I used the **LEFT() and RIGHT() formulas** instead.
* Created a new **Product Brand** column by combining the Product Name and Brand Name using:
  `=CONCATENATE([@[Product Name]]," ",[@[Brand Name]])`

### 4. Number Formatting

* Formatted the **Price** column as **Currency**.
* Tried to format the **Manufacturing Date** column as **DD-MM-YYYY**.
* The date format did not change as expected, even though Excel displayed the values as dates in the formula bar.

### 5. Conditional Formatting

* Applied **Data Bars** to the **Price** column to visually compare the prices.
* Created a **custom Conditional Formatting rule** to highlight cells where the **Category is Electronics**.
