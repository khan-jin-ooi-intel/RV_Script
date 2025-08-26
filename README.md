### <ins>Description</ins>
- To pull preliminary unit data for reject validation purposes

### <ins>Prerequisites</ins> 
- `Python`  <sub>( ... require user installation )</sub>
- `Pandas` <sub>( ... automatically installed with use of batch file )</sub>

### <ins>How to Use</ins>
1) Save all 3 files into your desired path <sup>( python, batch and xlsx )</sup>
2) Configure the Format File <sup>(xlsx)</sup> accordingly to the [steps here](ConfiguringTheFormatFile.md)   (First time users will need to do this, others can skip)
3) Configure and Run the Batch file (Specifically on the "input arguments" section)
<img width="550" height="140" alt="image" src="https://github.com/user-attachments/assets/bc955576-cd3c-41e9-b807-01377f9eb7c2" />

### <ins>Using python script individually<ins>
- reccommended to use run using provided Batch File 

| Flag | Required / Optional | Remark |
| - | - | - |
| --inputfile | Required | path to input file (raw data file pulled from Aqua in csv format) |
| --outputfile | Required | path to place output file |
| --format | Required | path to `'format'` file |
| --vid | Required | list of comma-seperated VIDs to pull for |
| --locn | Required | list of comma-seperated locations/sockets to pull for | 
| --dump | Optional | Enabling this flag will print out all corresponding values for all defined tokens in a seperate sheet | 

- example cmd line
  
<pre><code>python RV_prelimauto_v2.4.1.py --inputfile "input file path" --outputfile "output file path" --format "format file path" --vid "VID1,VID2,VID3,VID4" --locn "6261,6212,5242,5243"</code></pre>

### <ins>Self-Help</ins>
| Problem | Example | Solution |
| - | - | - |
| 1. My output table shows "Duplicates Found, Please Optimize Keywords", how can I resolve this? | <img width="338" height="181" alt="image" src="https://github.com/user-attachments/assets/a7b892d8-4eb2-4d5d-80ec-9c0c0360acef" /> | console will point to the token (blue) and all matches (red) based on the "Keywords" & "Exclude_Keywords" defined in the Format File, thus optimize your parameters accordingly  <img width="951" height="151" alt="image" src="https://github.com/user-attachments/assets/dae2441a-4b2f-4ab3-9753-5e91600deac1" /> |
| 2. The raw data (csv) file does not contain the data I'm looking for, what can I do? | - | Modify this line in the batch file to point to your customized aqua report <img width="1700" height="25" alt="image" src="https://github.com/user-attachments/assets/a5fc9d79-9832-482b-b28b-6b783eb384e7" /> |




