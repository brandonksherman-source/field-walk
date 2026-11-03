# Field Walk

Offline site-walk app for an iPad. Fill in the building, add walk-through notes, take photos, then **Send to estimator** to save one `field-walk-….json` file and open it in the Photo Estimator → *Import from field*.

- Live at https://brandonksherman-source.github.io/field-walk/
- On the iPad: open in Safari → Share → **Add to Home Screen**, then always open it from that icon. Works with no signal after the first load.
- Site walks and photos are stored only on the device (IndexedDB). Nothing is uploaded from the app; this repo holds only the app's code.
- Updating: edit the files, bump `VERSION` in `sw.js`, commit.
