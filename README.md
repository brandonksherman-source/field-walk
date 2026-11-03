# Field Walk

Offline iPad app for Ahrens Contracting site walks. Fill in the building, take photos, then **Export** a `field-walk-….json` file and open it in the Ahrens **Photo Estimator** → *Import from field*.

- Live at https://brandonksherman-source.github.io/field-walk/
- On the iPad: open in Safari → Share → **Add to Home Screen**. Works with no signal after the first load.
- Everything (walks, photos) is stored only on the device in IndexedDB. Nothing is uploaded.
- Updating: edit `index.html`, bump `VERSION` in `sw.js`, commit.
