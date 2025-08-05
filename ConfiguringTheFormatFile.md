Referring to the `format.xlsx` file as reference:

## "sort_tokens"

| Parameter | Description |
|-|-|
| Variable | self-defined token name to be used later |
| Default_Value | dummy placeholder value if data does not exist |
| Keywords | comma-seperated words to match the "TEST_NAME" you're trying to pull |
| Exclude_Keywords | comma-seperated words to exclude from the "TEST_NAME" you're trying to pull |

Notes:
- Compulsory to create a `"_first"` & `"_retest"` token for each data (example below) 

<img width="761" height="120" alt="image" src="https://github.com/user-attachments/assets/beab0783-c25b-44e3-9578-7145acd262a4" />

-------------------------------------------------------------------------------------------------------------------------------------
## "class_tokens"
Notes:
- the same is applied as in the `"sort_tokens"` description
- Define tokens based on `"TEST_NAME"`, not on which socket (6261, 6212, 6242, etc.)
  - ex. to pull for `"SCN_GT::CTRL_X_GFXAGG_K_POSTHVQK_X_X_X_X_GT_EVALUATE_SKU"` for both 6261 & 6212 will only require a single token defined
-------------------------------------------------------------------------------------------------------------------------------------
## "sample"
-------------------------------------------------------------------------------------------------------------------------------------
## "table_params"
-------------------------------------------------------------------------------------------------------------------------------------
## "compare"
