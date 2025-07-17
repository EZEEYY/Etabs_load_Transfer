# 🔁 ETABS Support Reaction Transfer Tool

This Python-based tool extracts **support reactions** from one ETABS `.EDB` model and applies them as **joint loads** in another model. It uses the **ETABS COM API** for automation and provides a **GUI interface** for user interaction.

## 📥 Download

👉 [Download etabs_transfer.exe](./dist/ETABS_LOAD_TRANSFER.exe)

Points to consider:
- Make sure both etabs files have same load pattern defined.
- This program only applies joints loads for cases which are defined in load pattern, so although in excel output we can get reactions for modal, response load,etc. they cannot be applied to the destination etabs file.
- Make sure the destination etabs file is unlocked.