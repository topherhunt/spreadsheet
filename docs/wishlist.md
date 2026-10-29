# Wish list

## Windows release

The NSIS target in `package.json` and the **Build** and **Windows Tests** workflows exist, but neither workflow has run yet. To ship:

- Run both workflows and fix whatever fails on Windows, then smoke-test the `.exe` on a real machine.
- Attach the `.exe` from the Build workflow to each release, since `scripts/release.sh` only uploads the `.dmg`.
- Add Windows install steps to the README. The installer is unsigned, so SmartScreen blocks it until you click **More info > Run anyway**.
- Mention Windows again in the README's opening line.
