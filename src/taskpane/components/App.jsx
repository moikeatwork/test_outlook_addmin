import * as React from "react";
import { useState, useEffect } from "react";
import PropTypes from "prop-types";
import { 
  Button, 
  Field, 
  Combobox,
  Option,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  tokens, 
  makeStyles 
} from "@fluentui/react-components";
import { 
  CheckmarkCircle24Regular, 
  ErrorCircle24Regular,
  ShieldCheckmark24Regular
} from "@fluentui/react-icons";
import { archiveEmail, searchAccounts } from "../taskpane";
import { authService } from "../authService";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "24px",
    display: "flex",
    flexDirection: "column",
  },
  header: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "8px",
    color: tokens.colorBrandForeground1,
    lineHeight: "1.2",
  },
  authBadge: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "6px 10px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    marginBottom: "24px",
    alignSelf: "flex-start",
  },
  form: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    flex: 1,
  },
  submitButton: {
    marginTop: "4px",
  },
  messageBar: {
    marginTop: "20px",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  
  const [userEmail, setUserEmail] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [searchResults, setSearchResults] = useState([]);
  const [selectedAccount, setSelectedAccount] = useState(null);
  const [isSearching, setIsSearching] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [result, setResult] = useState(null);

  // Get user email on mount
  useEffect(() => {
    if (typeof Office !== 'undefined' && Office.context?.mailbox?.userProfile) {
      setUserEmail(Office.context.mailbox.userProfile.emailAddress);
    }
  }, []);

  // Debounced search
  useEffect(() => {
    if (searchQuery.length < 3) {
      setSearchResults([]);
      return;
    }

    const timeoutId = setTimeout(async () => {
      setIsSearching(true);
      const response = await searchAccounts(searchQuery);
      setIsSearching(false);
      
      if (response.success) {
        setSearchResults(response.results);
      } else {
        setResult({ success: false, error: response.error });
      }
    }, 300);

    return () => clearTimeout(timeoutId);
  }, [searchQuery]);

  const handleAccountSelect = (_, data) => {
    const account = searchResults.find(acc => acc.id === data.optionValue);
    setSelectedAccount(account);
    setSearchQuery(account?.name || "");
  };

  const handleSubmit = async () => {
    if (!selectedAccount) {
      setResult({ success: false, error: "Please select an account" });
      return;
    }

    setIsSubmitting(true);
    setResult(null);

    const response = await archiveEmail(selectedAccount.id, selectedAccount.name);
    
    setIsSubmitting(false);
    setResult(response.success 
      ? { success: true, message: "Email archived to SugarCRM successfully" }
      : { success: false, error: response.error }
    );
    
    if (response.success) {
      setSearchQuery("");
      setSearchResults([]);
      setSelectedAccount(null);
    }
  };

  return (
    <div className={styles.root}>
      <h1 className={styles.header}>Archive Email</h1>
      
      <div className={styles.authBadge}>
        <ShieldCheckmark24Regular />
        <span>{userEmail}</span>
      </div>

      <div className={styles.form}>
        <Field 
          label="Account"
          hint="Type to search CRM accounts"
        >
          <Combobox
            placeholder="Start typing account name..."
            value={searchQuery}
            onInput={(e) => setSearchQuery(e.target.value)}
            onOptionSelect={handleAccountSelect}
            disabled={isSubmitting}
          >
            {isSearching && (
              <Option key="searching" disabled>
                <Spinner size="tiny" /> Searching...
              </Option>
            )}
            {!isSearching && searchResults.length === 0 && searchQuery.length >= 2 && (
              <Option key="no-results" disabled>
                No accounts found
              </Option>
            )}
            {!isSearching && searchResults.map((account) => (
              <Option key={account.id} value={account.id} text={account.name}>
                {account.name}
              </Option>
            ))}
          </Combobox>
        </Field>

        <Button
          appearance="primary"
          size="large"
          onClick={handleSubmit}
          disabled={isSubmitting || !selectedAccount}
          className={styles.submitButton}
          icon={isSubmitting ? <Spinner size="tiny" /> : undefined}
        >
          {isSubmitting ? "Archiving..." : "Archive Email"}
        </Button>

        {result && (
          <MessageBar
            intent={result.success ? "success" : "error"}
            className={styles.messageBar}
            icon={result.success ? <CheckmarkCircle24Regular /> : <ErrorCircle24Regular />}
          >
            <MessageBarBody>
              <MessageBarTitle>
                {result.success ? "Success" : "Error"}
              </MessageBarTitle>
              {result.message || result.error || "An error occurred"}
            </MessageBarBody>
          </MessageBar>
        )}
      </div>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;