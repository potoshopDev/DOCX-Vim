# Add project specific ProGuard rules here.
# Rust calls these methods through JNI by their exact Java names and signatures.
# R8 cannot see those call sites, so release builds must retain the methods.
-keepclassmembers class ru.docxvim.app.MainActivity {
    public void registerDocxTask(java.lang.String, java.lang.String);
    public void focusDocxTask(java.lang.String, java.lang.String);
    public void closeDocxTask(java.lang.String, java.lang.String);
}
# You can control the set of applied configuration files using the
# proguardFiles setting in build.gradle.
#
# For more details, see
#   http://developer.android.com/guide/developing/tools/proguard.html

# If your project uses WebView with JS, uncomment the following
# and specify the fully qualified class name to the JavaScript interface
# class:
#-keepclassmembers class fqcn.of.javascript.interface.for.webview {
#   public *;
#}

# Uncomment this to preserve the line number information for
# debugging stack traces.
#-keepattributes SourceFile,LineNumberTable

# If you keep the line number information, uncomment this to
# hide the original source file name.
#-renamesourcefileattribute SourceFile
