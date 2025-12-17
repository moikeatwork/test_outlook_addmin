/* global Office console */

export async function archiveEmail(identifier, identifierType) {
  try {
    // Get email metadata
    const item = Office.context.mailbox.item;
    const userEmail = Office.context.mailbox.userProfile.emailAddress;
    
    if (!item) {
      throw new Error("No email item found");
    }

    // Get the email ID (EWS ID format)
    const messageId = Office.context.mailbox.item.itemId;
    
    if (!messageId) {
      throw new Error("Unable to get message ID");
    }

    // Prepare webhook payload
    const payload = {
      messageId: messageId,
      userPrincipalName: userEmail,
      identifierType: identifierType, // "domain" or "accountName"
      identifier: identifier
    };

    // Call N8N webhook
    const response = await fetch("https://workflows.prostarpics.com/webhook-test/archive_sugarcrm", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      throw new Error(`Webhook failed: ${response.status} ${response.statusText}`);
    }

    const result = await response.json();
    return { success: true, data: result };

  } catch (error) {
    console.error("Archive error:", error);
    return { success: false, error: error.message };
  }
}