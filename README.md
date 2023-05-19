# GPT for Excel
This macro-enabled XSLM workbook allows for direct calls to the [OpenAI GPT API](https://openai.com/product#made-for-developers) via two inline functions, GPT() and GPTFillRange(). 
### Please note: This XLSM workbook only works on Windows!

## Prerequisites
For the macros on this worksheet to function properly, please follow the steps outlined below.

### Obtain an OpenAI API Key
You must obtain your own OpenAI API key for the macros in this workbook to function properly.
For instructions on procuring an OpenAI API key, please see 
https://www.windowscentral.com/software-apps/how-to-get-an-openai-api-key

After receiving the OpenAI API key, insert it in the API Key cell on the "Configuration" tab
!['API key'](/assets/001-enter-api-key.png)

The API Key field will change background color to blue
!['API key entered'](/assets/002-enter-api-key.png)

### Log data
If you'd like to keep an eye on costs, set "Log Data" to "Yes" in the "Configuration" tab. This will trigger the generation of a text file called "GPTForExcelLog.txt" in the same folder as the macro workbook to log the time, prompt, completion, total tokens and cost of each GPT API call. 

!['Log data'](/assets/015-log-data.png)

### GPTFillRange() Prompt Configuration for One-shot / Few-shot / Multi-shot learning
Modify the values in this section as needed. The default values perform quite well for most one-shot/few-shot/multi-shot learning tasks, so please only make changes to these values if the defaults do not deliver the desired results. 

### "#Value!" results
During times of high demand, individual calls to the OpenAI API may time out. If this happens, the macro will simply continue processing the next operation and will display "#Value!" in the impacted cell. To clear this error, re-run the operation in the cell showing the "#Value!" result.

### Enable macros
Depending on the delivery mode, you may have to accept one or more warnings about potentially unsafe content, i.e. the macro code in this workbook 

#### If you encounter the Microsoft Excel Security Notice
Select "Enable Macros"

!['Enable macros'](/assets/003-ms-security-warning.png)

#### When opening the workbook
Select "Enable Content"

!['Enable content'](/assets/004-enable-content.png)

#### If you encounter the "Microsoft has blocked macros" issue
!['Enable content'](/assets/005-ms-blocked.png)

Select the properties of the xlsm file and select "Unblock"

!['Enable content'](/assets/006-ms-unblock.png)

## Special note about Automatic Calculation of formulas
As the GPT() and GPTFillRange() functions are "auto-calculated" during many common spreadsheet activities, you may want to disable Automatic Calculations in the Formulas ribbon to avoid potentially costly repeated OpenAI API calls.

!['Automatic Calculation'](/assets/013-automatic-calculation.png)

!['Manual Calculation'](/assets/014-manual-calculation.png)


## GPT()
This function simply submits your prompt to GPT and returns the result based on the wording of the input prompt and the temperature (i.e. the amount of creativity)

### Basic use of GPT()
Create a new worksheet. Enter the desired prompt text into a cell. Optionally, you can enter the text directly into the cell by typing it into the formula like =GPT("Here is my prompt"). I recommend keeping the text in a separate cell as shown below. 

!['Basic example of GPT()'](/assets/007-gpt-simple.png)

Upon hitting "Enter," the prompt is submitted to GPT. Depending on system load and complexity of the task, response times may vary from several seconds to up to a minute. <b> You may see a "Not responding" message, but the system is still processing the request. </b>

!['Status message for GPT()'](/assets/008-gpt-dialog.png)

Once the result is returned from GPT, the output is displayed in the cell containing the formula.
!['Basic GPT() output'](/assets/009-gpt-simple-complete.png)

### Advanced use of GPT()
To create bespoke GPT results for a variety of inputs, arrange individual items into a table and use either & syntax or the CONCAT() function to create the desired input prompt for GPT()
!['Advanced GPT() setup'](/assets/010-gpt-simple-concat.png)

By filling the GPT() formula down the "Message" column, GPT creates a bespoke message for the respective inputs on each row.
!['Advanced GPT() results'](/assets/011-gpt-simple-concat-complete.png)

## GPTFillRange()
Designed for "few-shot learning," this function allows you to easily submit prompt/response pairs to GPT to customize completions.

#### Training Range:
Contains prompt completions examples - can be simple completions, token extractions, sentiment analysis, etc.

#### Input Range:
Contains prompt inputs to be completed by GPT.

#### Fill Range:
Will be automatically populated by GPTFillRange()

!['GPTFillRange()'](/assets/012-fill-range-01.png)

### Step-by-Step Instructions - Basic Two Column Example

In the first empty cell immediately following the Training Range and immediately to the right of the Input Range, type GPTFillRange()
!['GPTFillRange()'](/assets/012-fill-range-02.png)

Next, select the Training Range
!['GPTFillRange()'](/assets/012-fill-range-03.png)

Type  comma, and then select the Input Range
!['GPTFillRange()'](/assets/012-fill-range-04.png)

Now complete the formula by typing a closing parenthesis. This will launch the GPT API calls and you will see a status dialog with the progress of the overall fill process. 
!['GPTFillRange()'](/assets/012-fill-range-05.png)

Upon successfully completing the fill process, the GPT-generated results will be displayed in the Fill Range.
!['GPTFillRange()'](/assets/012-fill-range-06.png)

### Step-by-Step Instructions - Advanced Multi-Column Example

Enter the GPTFillRange() command into the appropriate cell covering the multi-columnar training range and the desired InputRange
!['GPTFillRange()'](/assets/012-fill-range-08.png)

After processing completes, the corresponding extracted values will be inserted into the appropriate columns
!['GPTFillRange()'](/assets/012-fill-range-09.png)


## Stay connected!
For more articles, news and thoughts on AI, follow me on 

[!['LinkedIn'](/assets/linkedin.png)](https://www.linkedin.com/in/bjornaustraat/)


