CORRECTED HALL TICKET GENERATOR

Fixes in this version:
1. PDF no longer displays literal <br> tags.
2. PDF shows "Jr. Supervisor" and "Signature" on separate lines.
3. Photo placeholder is exactly "(Paste Photograph Here)" when no photo is supplied.
4. Word template removes literal <br> tags.
5. Word output is compacted and the extra Academic Year paragraph was removed so the Hall Ticket remains on one page.
6. Both Word and PDF are generated for every student.
7. ZIP contains separate Word_Hall_Tickets and PDF_Hall_Tickets folders.

Run:
pip install -r requirements.txt
streamlit run app.py
