# StadeControl Android wrapper

This project creates a simple Android application that opens the Railway-hosted StadeControl app in a WebView.

## Important

Update the URL in `MainActivity.kt` to your exact Railway domain if needed.

Default URL:

`https://stadecontrol-production.up.railway.app`

## Build

1. Install Android Studio and the Android SDK.
2. Open this project in Android Studio.
3. Sync Gradle.
4. Choose `Build > Build Bundle(s) / APK(s) > Build APK`.

## Notes

- This wrapper is designed to open an existing web app as a mobile shell.
- The app is intentionally lightweight and uses a single `WebView` for the user experience.
