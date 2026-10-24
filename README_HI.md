# Bank E-Auction Extractor — Cloud setup

## GitHub par upload
Repository ke root mein ye 4 files upload karein:
- `app.py`
- `requirements.txt`
- `packages.txt`
- `README_HI.md`

ZIP ko repository mein upload karna kaafi nahi; pehle ZIP extract karke andar ki files upload karein.

## Streamlit Community Cloud deploy
1. https://share.streamlit.io par GitHub se sign in karein.
2. Create app → repository `bank-eauction-extractor` select karein.
3. Branch `main`, Main file path `app.py`.
4. Deploy par click karein.

## Important
- Community Cloud deployment par uploaded notice images server par process hoti hain. Sensitive borrower data ke liye public app use na karein.
- Private app availability/limits account ke mutabiq ho sakte hain. Agar app private rakhna zaroori hai, app ki Sharing settings check karein.
- OCR galti kar sakta hai; borrower (co-borrower nahi), amounts, dates, times aur property description original notice se verify karein.
- Streamlit Cloud free resources/availability badal sakte hain. Agar build error aaye to Manage app → Logs dekhein.
