The files in this folder were gathered by Prof Engelmann on his fieldtrip to watch Stacey's process. Here is the email he sent explaining them: 

"Attached you get an anonymised excerpt of the relevant files. Export_1 and 2 are the magic pdf input.
Until there is a data share agreement in place, my suggestion is to mock the PDF parsing with just manually crafting a json file and using this as input.
Make the system modular to easily attach the PDF parsing part, as it should be flexible to accommodate other school report formats.

The part_data xlsx shows exactly the same data in the excel sheet. Each week the sheet grows by a column. The colour coding easily visualises the onset of the threshold missed hours (Which is the sum of Excused+Unexcused Hours, but not Medial or suspension).

The Custodian and Address currently is manually asked for from the school. A data reduction would be to only ask for the addresses if required because of a threshold requiring to send a letter.
The last column is more of a timeline to keep track of dates for predefined events (invitation letter sent) and other unstructured information.
It is always a list of: Date + Text.
Given the development of the missing hours, sometimes the letter is delayed by a week to allow for race conditions with handing in medial excuses.
Other times parents opt out of mediation and don't want to be contacted again.

On the first threshold, the preliminary letter should be generated. On the second threshold, it is the Invitation letter (docx) and the mediation agreement form, which does not yet include any templatable data.
Along with all letters, the system also needs a convenient way to print addresses onto envelopes.
To get all the information for the invitation, a empty slot on a Friday needs to be allocated.
This is tracked in the scheduling tab. Mostly by appending the meeting at the earliest free spot, but at least 3 weeks in advance.
After the mediation, this table also tracks the outcome reached with the parents. (Meaning if the parent attended and accepted mediation, not if the truancy hours were reduced.)"

