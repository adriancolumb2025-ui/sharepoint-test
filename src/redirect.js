// This dedicated page MUST contain only MSAL's redirect bridge.
// Do not load third-party scripts, analytics or application code here.
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

broadcastResponseToMainFrame();
