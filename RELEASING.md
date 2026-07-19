# Releasing kpoi

## One-time setup

1. A [central.sonatype.com](https://central.sonatype.com) account with the
   `io.github.kouroshmsv` namespace verified (sign in with GitHub).
2. A GnuPG key whose public part is on `keyserver.ubuntu.com`.
3. In `$GRADLE_USER_HOME/gradle.properties` (defaults to `~/.gradle`; never commit these):

   ```properties
   mavenCentralUsername=<portal user token name>
   mavenCentralPassword=<portal user token password>
   signing.gnupg.keyName=<gpg key id>
   # Windows, when gpg is not on the daemon's PATH:
   signing.gnupg.executable=C:/Program Files/GnuPG/bin/gpg.exe
   ```

   Generate the token at central.sonatype.com → account icon → *Generate User Token*.

## Per release

1. Set `version` in the root `build.gradle.kts` (drop `-SNAPSHOT`).
2. `./gradlew clean build`
3. `./gradlew publishToMavenCentral` — signs and uploads a deployment bundle
   (gpg will prompt for the key passphrase).
4. Review the deployment at
   [central.sonatype.com/publishing](https://central.sonatype.com/publishing/deployments)
   and press **Publish**. Propagation to Maven Central takes ~15–60 minutes.
5. Tag and push: `git tag vX.Y.Z && git push origin vX.Y.Z`, then create a
   GitHub release from the tag.
6. Bump `version` to the next `X.Y.Z-SNAPSHOT` and commit.
