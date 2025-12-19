/* global Office console */

import { authService } from './authService';

export async function searchAccounts(query) {
  try {
    if (query.length < 2) {
      return { success: true, results: [] };
    }

    const userEmail = Office.context.mailbox.userProfile.emailAddress;

    const response = await authService.makeAuthenticatedRequest('/search_accounts', {
      method: 'POST',
      body: JSON.stringify({ 
        query: query,
        userEmail: userEmail
      })
    });

    if (!response.ok) {
      throw new Error(`Search failed: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    return { success: true, results: data.results || [] };

  } catch (error) {
    console.error("Search error:", error);
    return { success: false, error: error.message, results: [] };
  }
}

export async function archiveEmail(accountId, accountName) {
  try {
    const item = Office.context.mailbox.item;
    const userEmail = Office.context.mailbox.userProfile.emailAddress;
    
    if (!item) {
      throw new Error("No email item found");
    }

    const messageId = item.itemId;
    
    if (!messageId) {
      throw new Error("Unable to get message ID");
    }

    const payload = {
      messageId: messageId,
      userPrincipalName: userEmail,
      accountId: accountId,
      accountName: accountName
    };

    const response = await authService.makeAuthenticatedRequest('/archive_sugarcrm', {
      method: 'POST',
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      throw new Error(`Archive failed: ${response.status} ${response.statusText}`);
    }

    const result = await response.json();
    return { success: true, data: result };

  } catch (error) {
    console.error("Archive error:", error);
    return { success: false, error: error.message };
  }
}