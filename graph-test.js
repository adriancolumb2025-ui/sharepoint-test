console.log("Testing Microsoft Graph developer environment");

async function testGraphConnection() {

    console.log("Simulating Microsoft Graph API request...");

    const mockGraphResponse = {
        user: "developer@example.com",
        displayName: "Microsoft Graph Test User",
        tenant: "Microsoft 365 Dev Sandbox"
    };

    console.log("Graph response:", mockGraphResponse);

    alert("Graph API test simulated. Check the console.");
}

testGraphConnection();
