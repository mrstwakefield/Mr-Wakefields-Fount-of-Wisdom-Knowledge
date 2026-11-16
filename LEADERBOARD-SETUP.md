# Setting up the class leaderboard

The leaderboard is already built into the site. It stays hidden until it's connected to a free Firebase project that you own. This takes about 15 minutes: part A happens in Firebase, part B in Microsoft Azure (with your school account), and part C links the two.

**What gets stored:** each student's first name and surname initial (e.g. "Aisha K"), their XP and their level. Firebase's sign-in system also keeps the email address and display name of everyone who signs in, which only you can see in the Firebase console. Everything else (cards, scrolls, themes) stays on the student's own device, as now.

---

## A. Firebase

1. Go to **console.firebase.google.com** and sign in with a Google account (yours or a school one).
2. Click **Create a project**. Name it something like `fount-leaderboard`. Turn **Google Analytics off**. Click **Create project**.
3. In the left menu: **Build → Firestore Database → Create database**.
   - Choose **Start in production mode**.
   - Location: **europe-west2 (London)**.
   - Click **Create**.
4. Click the **⚙️ gear → Project settings**. Under **Your apps**, click the **`</>`** (Web) icon.
   - App nickname: `Fount`. Don't tick Firebase Hosting.
   - Click **Register app**. You'll see a block of code containing `firebaseConfig = { apiKey: …, authDomain: …, projectId: … }`.
   - **Copy that block**: you'll send it to me (it's safe to share; it isn't a password).
5. In the left menu: **Build → Authentication → Get started → Sign-in method → Add new provider → Microsoft**.
   - Turn it on. **Copy the redirect URL** it shows (it ends in `/__/auth/handler`).
   - Leave this tab open: you need to paste two things into it in part C.

## B. Microsoft (Azure)

6. Go to **portal.azure.com** and sign in with your **school** Microsoft account.
7. Search for **App registrations** and click **New registration**.
   - Name: `Fount of Wisdom leaderboard`.
   - Supported account types: **Accounts in this organizational directory only**.
   - Redirect URI: choose **Web** and paste the URL from step 5.
   - Click **Register**.
8. On the app's **Overview** page, copy two things:
   - the **Application (client) ID**
   - the **Directory (tenant) ID**
9. Go to **Certificates & secrets → New client secret**. Add it (24 months is fine), then **copy the Value** straight away, because it's only shown once.

> If Azure says you aren't allowed to register apps, or students later see "Need admin approval", forward this guide to IT. They need to register the app, or grant admin consent for the **User.Read** permission. It's a two-minute job for them.

## C. Link them up

10. Back in Firebase (the Microsoft provider from step 5):
    - paste the **Application (client) ID** and the **client secret Value**
    - click **Save**
11. In Firebase: **Authentication → Settings → Authorized domains → Add domain**, then type `mrstwakefield.github.io`.
12. **Send me these three things:**
    - the `firebaseConfig` block from step 4
    - the **Directory (tenant) ID** from step 8
    - your school email domain: the part after the @, e.g. `myschool.org.uk`

    **Don't send the client secret.** It only goes into Firebase.

I'll then switch the leaderboard on and give you the security rules to paste into **Firestore → Rules**. Those rules make sure only school accounts can see it, and that each student can only update their own entry.
