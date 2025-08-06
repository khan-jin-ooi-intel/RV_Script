Referring to the `format.xlsx` file as reference:

## "sort_tokens"
- Compulsory to create a `"_first"` & `"_retest"` token for each data (example below)
  
| Parameter | Description |
|-|-|
| Variable | self-defined token name to be used later |
| Default_Value | dummy placeholder value if data does not exist |
| Keywords | comma-seperated words to match the "TEST_NAME" you're trying to pull |
| Exclude_Keywords | comma-seperated words to exclude from the "TEST_NAME" you're trying to pull |

<img width="500" height="120" alt="image" src="https://github.com/user-attachments/assets/beab0783-c25b-44e3-9578-7145acd262a4" />

<sub>Figure 1. Params Example</sub>

---
## "class_tokens"
- Same is applied as in the description for `"sort_tokens"` 
- Define tokens based on `"TEST_NAME"`, not on which socket (6261, 6212, 6242, etc.)
  - eg. to pull for `"SCN_GT::GTMISC_X_PRIME_K_BEGIN_GPU_X_X_X_AGG_EVALUATE_SKU"` for both 6261 & 6212 will only require a single token defined
---
## "sample"
- Define output display format 
  - Use defined tokens with socket as suffix (ex. `"sort_"`, `"classhot_"`, `"qahot_"`) for Search & Replace

<img width="1641" height="220" alt="image" src="https://github.com/user-attachments/assets/a120d576-7789-4644-a936-ecef8b24bbb5" />

<sub>Figure 2. Sample Output</sub>

---
## "table_params"
- Insert the range of tokens defined in `sample` sheet for Search & Replace operation
  - Using Figure 2 as example, `Columns` = "B:F", `StartRow` = 3, `EndRow` = 7.
---
## "compare"
- Optional function to compare and highlight anomalies
- Define the cells for comparison and the odd values will be highlighted

|Comparison|<img width="1030" height="60" alt="image" src="https://github.com/user-attachments/assets/7b234ba5-5f4e-4bbc-a900-5801f4e4d279" />|
|-|-|
|Result|<img width="1590" height="307" alt="image" src="https://github.com/user-attachments/assets/b0fe5a53-f633-4e2e-bc76-fe82da4d440b" />|






