import java.io.FileInputStream
import java.util.Properties

plugins {
    id("com.android.application")
    id("org.jetbrains.kotlin.android")
}

android {
    namespace = "com.zys.excel2image"
    compileSdk = 34

    defaultConfig {
        applicationId = "com.zys.excel2image"
        minSdk = 23
        targetSdk = 34

        versionCode = 1
        versionName = "0.1.0"

        // Apache POI + dependencies can easily exceed 64K methods.
        multiDexEnabled = true
    }

    signingConfigs {
        create("release") {
            // For small-team side-load distribution:
            // - Put secrets in keystore.properties (NOT committed)
            // - If missing, fall back to debug signing so `assembleRelease` still works locally.
            val propsFile = rootProject.file("keystore.properties")
            if (propsFile.exists()) {
                val props = Properties().apply { load(FileInputStream(propsFile)) }
                storeFile = file(props.getProperty("storeFile"))
                storePassword = props.getProperty("storePassword")
                keyAlias = props.getProperty("keyAlias")
                keyPassword = props.getProperty("keyPassword")
            } else {
                initWith(getByName("debug"))
            }
        }
    }

    buildTypes {
        release {
            // Keep it simple for the first public test builds.
            // Apache POI is large and uses some reflection; minification can break things.
            isMinifyEnabled = false
            signingConfig = signingConfigs.getByName("release")
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro",
            )
        }
        debug {
            isMinifyEnabled = false
        }
    }

    buildFeatures {
        viewBinding = true
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }

    kotlinOptions {
        jvmTarget = "17"
    }

    packaging {
        resources {
            excludes += setOf(
                "META-INF/DEPENDENCIES",
                "META-INF/INDEX.LIST",
                "META-INF/LICENSE",
                "META-INF/LICENSE.txt",
                "META-INF/license.txt",
                "META-INF/NOTICE",
                "META-INF/NOTICE.txt",
                "META-INF/notice.txt",
                "META-INF/ASL2.0",
            )
        }
    }
}

dependencies {
    implementation("androidx.core:core-ktx:1.12.0")
    implementation("androidx.appcompat:appcompat:1.6.1")
    implementation("com.google.android.material:material:1.11.0")
    implementation("androidx.activity:activity-ktx:1.8.2")
    implementation("androidx.lifecycle:lifecycle-runtime-ktx:2.6.2")
    implementation("androidx.constraintlayout:constraintlayout:2.1.4")
    implementation("androidx.multidex:multidex:2.0.1")

    // Pinch-to-zoom preview (so wide tables can stay as a single image).
    implementation("com.github.chrisbanes:PhotoView:2.3.0")

    // Excel parsing (primarily .xlsx).
    implementation("org.apache.poi:poi-ooxml:5.2.5")
    implementation("javax.xml.stream:stax-api:1.0-2")
    // StAX implementation (Android doesn't ship javax.xml.stream).
    implementation("com.fasterxml:aalto-xml:1.3.3")
}
