# VidyaPath

VidyaPath is a study companion with source-grounded tutoring, course quizzes, and learner progress. Accounts use Supabase Auth; course files and quiz records are intended to be scoped to the signed-in user.

## Local setup

1. Create a Supabase project.
2. In the Supabase SQL Editor, run `supabase_setup.sql`. It creates quiz, video-source, and tutor-history tables with row-level security, plus a private `course-files` storage bucket with per-user path policies.
3. In Supabase Auth URL settings, set the local Site URL to `http://localhost:8501`. Add both `http://localhost:8501` and `http://localhost:8501/**` to allowed redirect URLs so local email confirmation, Google sign-in, and password-reset callbacks can return to the app.
4. Email/password registration is enabled through Supabase Auth. Keep email confirmation enabled for real users.
5. For Google sign-in, create a Google OAuth web client and enable Google in Supabase Auth. Use the callback URL shown on Supabase's Google provider settings page in the Google OAuth client configuration. Allow `http://localhost:8501` as the local app redirect.
6. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml`, then fill in the Supabase project URL, Supabase anon/publishable key, app URL, and Gemini key. Never use a Supabase service-role/secret key in this app, and do not commit the real secrets file.
7. In VS Code's terminal, stop Streamlit with **Ctrl+C**, then run:

   ```powershell
   python -m pip install -r requirements.txt
   streamlit run app.py
   ```

Sign in with email/password or Google. Each learner's files go under their Supabase Auth user ID, and database/storage row-level security policies restrict access to that same user. Tutor and quiz requests send the question and matching excerpts to Gemini; original files are stored privately in Supabase Storage. Signed-in users can revisit their recent tutor questions, answers, and cited sources; guest history lasts for the current session. Learners can also save and remove personal notes linked to a page, slide, or video timestamp. After a quiz, learners can retry only the questions they missed. Scanned PDFs can be transcribed from the course library after the learner checks the consent box; that action sends the selected PDF (up to 50 MB) to Gemini for handwriting transcription. The Video lessons tab accepts public YouTube links and creates a timestamped topic outline. With consent, the video is analyzed by Gemini; signed-in outlines are saved in the account's private `video_sources` table. Guest outlines last for the current session. Voice mode previews the recording, rejects silence, and lets the learner review or correct the transcript before sending a question to the course tutor. Voice processing is sent to Gemini only after the learner consents and chooses the voice action; tutor replies can be read aloud.

If you already configured Supabase, run `supabase_video_sources.sql`, `supabase_tutor_history.sql`, and `supabase_study_notes.sql` in its SQL Editor to create the private tables and row-level security policies for video outlines, saved tutor history, and study notes. New projects can use the updated `supabase_setup.sql` instead.

## Security notes

- A system cannot be promised to be “100% secure.” This app uses Supabase Auth and row-level security as the access-control foundation; configuration and deployment still need review before real multi-user release.
- The app uses the Supabase anon/publishable key with the signed-in user's session. Never add the service-role key to app secrets or source code.
- Google sign-in requires Google OAuth credentials configured in the Supabase dashboard. The button will not work until that provider setup is complete.
- The old local SQLite progress file from the single-user prototype is not imported: it used manually entered names and cannot safely identify account owners.
