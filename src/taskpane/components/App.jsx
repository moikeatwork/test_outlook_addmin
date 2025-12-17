import * as React from "react";
import { useState } from "react";
import PropTypes from "prop-types";
import { 
  Button, 
  Field, 
  Input, 
  RadioGroup, 
  Radio,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  tokens, 
  makeStyles 
} from "@fluentui/react-components";
import { CheckmarkCircle24Regular, ErrorCircle24Regular } from "@fluentui/react-icons";
import { archiveEmail } from "../taskpane";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "20px",
  },
  header: {
    fontSize: tokens.fontSizeHero800,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "20px",
    color: tokens.colorBrandForeground1,
  },
  form: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
  },
  submitButton: {
    marginTop: "10px",
  },
  messageBar: {
    marginTop: "20px",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  
  const [identifierType, setIdentifierType] = useState("domain");
  const [identifier, setIdentifier] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [result, setResult] = useState(null);

  const handleSubmit = async () => {
    if (!identifier.trim()) {
      setResult({ success: false, error: "Please enter a value" });
      return;
    }

    setIsSubmitting(true);
    setResult(null);

    const response = await archiveEmail(identifier.trim(), identifierType);
    
    setIsSubmitting(false);
    setResult(response);
    
    if (response.success) {
      // Clear form on success
      setIdentifier("");
    }
  };

  return (
    <div className={styles.root}>
      <h1 className={styles.header}>Archive to SugarCRM</h1>
      
      <div className={styles.form}>
        <RadioGroup
          value={identifierType}
          onChange={(_, data) => setIdentifierType(data.value)}
        >
          <Radio value="domain" label="Domain" />
          <Radio value="accountName" label="Account Name" />
        </RadioGroup>

        <Field 
          label={identifierType === "domain" ? "Enter Domain" : "Enter Account Name"}
          hint={identifierType === "domain" ? "e.g., example.com" : "e.g., Acme Corporation"}
        >
          <Input
            value={identifier}
            onChange={(_, data) => setIdentifier(data.value)}
            placeholder={identifierType === "domain" ? "example.com" : "Account name"}
            disabled={isSubmitting}
          />
        </Field>

        <Button
          appearance="primary"
          size="large"
          onClick={handleSubmit}
          disabled={isSubmitting || !identifier.trim()}
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
              {result.success 
                ? "Email archived to SugarCRM successfully"
                : result.error || "Failed to archive email"
              }
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