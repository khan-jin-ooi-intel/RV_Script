Referring to the `format.xlsx` file as reference:

## "sort_tokens"

| Parameter | Description |
|-|-|
| Variable | self-defined token name to be used later |
| Default_Value | dummy placeholder value if data does not exist |
| Keywords | comma-seperated words to match the "TEST_NAME" you're trying to pull |
| Exclude_Keywords | comma-seperated words to exclude from the "TEST_NAME" you're trying to pull |

- Compulsory to create a `"_first"` & `"_retest"` token for each data (example below) 

<img width="459" height="120" alt="image" src="https://github.com/user-attachments/assets/beab0783-c25b-44e3-9578-7145acd262a4" />

---
## "class_tokens"
- The same is applied as in the description for `"sort_tokens"` 
- Define tokens based on `"TEST_NAME"`, not on which socket (6261, 6212, 6242, etc.)
  - eg. to pull for "SCN_GT::GTMISC_X_PRIME_K_BEGIN_GPU_X_X_X_AGG_EVALUATE_SKU" for both 6261 & 6212 will only require a single token defined
---
## "sample"
- Define output display format 
  - insert defined tokens for Search & Replace operation 
---
## "table_params"
- Insert the range of tokens defined in `sample` sheet for Search & Replace operation 
---
## "compare"
- optional function to compare and highlight anomalies
