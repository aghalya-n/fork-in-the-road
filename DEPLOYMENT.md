# Austin Lunch Club: free deployment

The free setup uses:

- **GitHub Pages** for `austin-lunch-club.html`
- **Google Sheets + Apps Script** for shared restaurants, lunch history, and the leaderboard
- Browser storage only for the current Friday's wheel entries

## 1. Create the shared Google Sheet backend

1. Create a blank Google Sheet named **Austin Lunch Club Data**.
2. Copy the Sheet ID from its URL:

   `https://docs.google.com/spreadsheets/d/SHEET_ID_IS_HERE/edit`

3. In the Sheet, open **Extensions → Apps Script**.
4. Replace the contents of `Code.gs` with the supplied [`Code.gs`](Code.gs) file.
5. Replace `PASTE_GOOGLE_SHEET_ID_HERE` near the top with your Sheet ID.
6. Click **Deploy → New deployment**.
7. Choose **Web app**.
8. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone
9. Authorize the script and complete the deployment.
10. Copy the web app URL ending in `/exec`.

Open `austin-lunch-club.html`, choose **Connect shared data**, paste the `/exec` URL, and select **Connect & migrate**. Existing local restaurants and lunch history will be copied into the Sheet.

If you later edit `Code.gs`, create a new deployment version from **Deploy → Manage deployments → Edit**. Keep using the same `/exec` URL.

## 2. Enable Google ratings with a 21-day refresh

Google requires billing to be enabled even when usage remains inside its monthly free allowance.

1. Open the [Google Cloud Console](https://console.cloud.google.com/) and create or select a project.
2. Attach a billing account.
3. Open **APIs & Services → Library** and enable **Places API (New)**.
4. Open **APIs & Services → Credentials → Create credentials → API key**.
5. Edit the key:
   - Keep **Application restrictions** set to **None** because Apps Script makes the server-side requests.
   - Under **API restrictions**, restrict it to **Places API (New)**.
6. Do not place this key in the HTML or GitHub repository.
7. Return to the Apps Script project and open **Project Settings → Script properties**.
8. Add this property:

   - Property: `GOOGLE_PLACES_API_KEY`
   - Value: your API key

9. Replace the deployed `Code.gs` with the supplied version and create a new deployment version.
10. In the Apps Script editor, select `refreshGoogleRatings` and click **Run** once. Approve its permissions when prompted.
11. Confirm the Restaurants sheet now contains Google ratings, review counts, Place IDs, links, and refresh timestamps.
12. Select `installRatingRefreshTrigger` and click **Run** once.

The trigger checks once per day but makes no Places API calls unless a rating is at least 21 days old. Place IDs are retained, so after the first matching pass it only requests current details. With roughly 80 restaurants, this averages about 116 rating refresh calls per month.

If Google matches the wrong restaurant, add its exact address on the Restaurant List tab, then clear that row's **Google Place ID** and **Rating Updated At** cells before running `refreshGoogleRatings` again.

Set a conservative Places API quota and a Cloud Billing budget alert as additional safeguards. Budget alerts notify you but do not themselves stop charges.

## 3. Publish the site with GitHub Pages

1. Create a free public GitHub repository, for example `austin-lunch-club`.
2. Rename `austin-lunch-club.html` to `index.html`.
3. Upload `index.html` to the repository's default branch.
4. Open **Settings → Pages**.
5. Under **Build and deployment**, choose:
   - **Source:** Deploy from a branch
   - **Branch:** `main`
   - **Folder:** `/ (root)`
6. Save. GitHub will show a public URL similar to:

   `https://YOUR-USERNAME.github.io/austin-lunch-club/`

Open that URL, connect the Apps Script backend once, and test adding a lunch.

## 4. Embed it in Google Sites

In Google Sites:

1. Choose **Insert → Embed → By URL**.
2. Paste the GitHub Pages URL.
3. Insert it, resize the frame, and publish the Google Site.

The Apps Script endpoint is intentionally accessible to anyone who knows its URL so the static site can save data without sign-in. Do not store confidential information in this Sheet. Keep a backup of the Sheet and avoid publishing the `/exec` URL separately.
