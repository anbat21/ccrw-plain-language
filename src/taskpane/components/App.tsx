/* global Word */
import * as React from "react";
import { 
  Button, Text, Field, Dropdown, Option, RadioGroup, Radio, 
  makeStyles, tokens, Spinner, Card, Textarea, Divider
} from "@fluentui/react-components";
import { DirectLine } from "botframework-directlinejs";

// Professional Styling: Purple and White Theme [cite: 1262, 1387]
const useStyles = makeStyles({
  container: { 
    padding: "20px", 
    display: "flex", 
    flexDirection: "column", 
    gap: "15px", 
    backgroundColor: "#FAF9FE", 
    minHeight: "100vh" 
  },
  setupCard: { 
    padding: "20px", 
    borderTop: "6px solid #6200EE", 
    boxShadow: tokens.shadow16 
  },
  primaryButton: { 
    backgroundColor: "#6200EE", 
    color: "white", 
    marginTop: "10px",
    ":hover": { backgroundColor: "#4B0082" },
    ":disabled": { backgroundColor: tokens.colorNeutralBackgroundDisabled }
  },
  headerText: { color: "#6200EE", marginBottom: "5px" },
  issueItem: { 
    borderLeft: "4px solid #9C27B0", 
    padding: "10px", 
    marginBottom: "10px", 
    backgroundColor: "white" 
  }
});

// Context Interface based on Section 10.5 of the Standard 
interface AudienceContext {
  primaryAudience: string;
  secondaryAudience: string;
  urgency: string;      // Mandatory (10.5 f)
  attitudes: string;    // Mandatory (10.5 a)
  context: string;      // Mandatory (10.5 b)
  format: string;       // Mandatory (10.5 c)
  informationNeed: string; // Mandatory (10.5 d)
  informationMostImportant: string; // Mandatory (10.5 d)
  informationMustUnderstand: string; // Mandatory (10.5 d)
  purpose: string;      // Mandatory (10.5 e)
  purposeUse: string;   // Mandatory (10.5 e)
}

export default function App() {
  const styles = useStyles();
  
  // State Management
  const [isReady, setIsReady] = React.useState(false);
  const [loading, setLoading] = React.useState(false);
  const [results, setResults] = React.useState<any>(null);
  const [status, setStatus] = React.useState("System Ready.");
  const [rawResponse, setRawResponse] = React.useState<string>("");
  const [skipped, setSkipped] = React.useState<{ text: string; reason: string }[]>([]);
  const [appliedCount, setAppliedCount] = React.useState(0);
  const [directLine, setDirectLine] = React.useState<any>(null);
  const [botStatus, setBotStatus] = React.useState<string>("Connecting to bot...");

  // Initialize DirectLine with secure token from server
  React.useEffect(() => {
    const startBot = async () => {
      try {
        // Call the internal API created in server.js
        const res = await fetch('/api/get-token');
        if (!res.ok) {
          setBotStatus(`Error: Token API failed (${res.status})`);
          setStatus("Bot connection failed. Check server and /api/get-token.");
          return;
        }
        const { token } = await res.json();
        if (!token) {
          setBotStatus("Error: No token received");
          setStatus("Bot connection failed. Token response was empty.");
          return;
        }
        
        // Initialize DirectLine using the temporary Token instead of the permanent Secret
        const dl = new DirectLine({ token: token });
        setDirectLine(dl);
        setBotStatus("Bot connected");
        setStatus("System Ready.");
      } catch (error) {
        console.error('Error initializing bot:', error);
        setBotStatus("Error: Bot connection failed");
        setStatus("Bot connection failed. Check server and secret configuration.");
      }
    };
    startBot();
  }, []);

  // Initialization Form Data [cite: 541, 602]
  const [aud, setAud] = React.useState<AudienceContext>({
    primaryAudience: "",
    secondaryAudience: "",
    urgency: "",
    attitudes: "",
    context: "",
    format: "",
    informationNeed: "",
    informationMostImportant: "",
    informationMustUnderstand: "",
    purpose: ""
    ,purposeUse: ""
  });

  // FORM VALIDATION: Check if any required field is empty [cite: 1510, 1537]
  const isFormValid = 
    aud.primaryAudience !== "" && 
    aud.secondaryAudience.trim().length > 0 &&
    aud.urgency !== "" && 
    aud.attitudes.trim().length > 5 && 
    aud.context !== "" && 
    aud.format !== "" &&
    aud.informationNeed.trim().length > 5 &&
    aud.informationMostImportant.trim().length > 5 &&
    aud.informationMustUnderstand.trim().length > 5 &&
    aud.purpose !== "" &&
    aud.purposeUse.trim().length > 5;

  const validatePlan = (data: any) => {
    if (!data || data.version !== "1.0" || data.scope !== "selection" || !Array.isArray(data.items) || !data.summary) {
      return { valid: false, reason: "Missing required top-level fields." };
    }
    if (typeof data.plainnessScore !== "number") {
      return { valid: false, reason: "Missing or invalid plainnessScore." };
    }
    const validItems = data.items.filter((item: any) => {
      return item.type === "highlight" &&
        item.match?.strategy === "exactText" &&
        typeof item.match?.text === "string" && item.match.text.trim().length > 0 &&
        item.style?.color &&
        typeof item.note?.label === "string" &&
        item.note?.message;
    });
    if (validItems.length !== data.items.length) {
      return { valid: false, reason: "One or more items are invalid." };
    }
    return { valid: true, reason: "" };
  };

  const analyzeSelection = async () => {
    if (!directLine) {
      setStatus("Bot not connected yet. Please wait...");
      return;
    }

    setLoading(true);
    setStatus("Processing...");
    setResults(null);
    setRawResponse("");
    setSkipped([]);
    setAppliedCount(0);
    
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          setStatus("Please select text in Word before analyzing.");
          setLoading(false);
          return;
        }

        // Feed context and text to the Brain
        const data: any = await new Promise((resolve, reject) => {
          const sub = directLine.activity$.subscribe(act => {
            if (act.from.role === "bot" && (act as any).text) {
              const match = (act as any).text.match(/\{[\s\S]*\}/);
              if (match) {
                try {
                  const parsed = JSON.parse(match[0]);
                  sub.unsubscribe();
                  resolve(parsed);
                } catch {
                  sub.unsubscribe();
                  reject(new Error("Bot response was not valid JSON."));
                }
              }
            }
          });

          const prompt = `You are the AI Analysis Engine for the CCRW Plain Language Reviewer.\n` +
            `Use ONLY the CAN-ASC-3.1:2025 document as your source.\n` +
            `Focus on: sentence length (max 12 words), passive voice, jargon/complex terminology.\n` +
            `Return ONLY a strict JSON object with this schema:\n` +
            `{"version":"1.0","scope":"selection","plainnessScore":0,"items":[{"id":"A1","type":"highlight","match":{"strategy":"exactText","text":"..."},"style":{"color":"yellow"},"note":{"label":"Plain language","message":"Replace with 'use' for clarity"}}],"summary":{"total":0,"categories":{}}}\n` +
            `Do NOT include any extra text or markdown.\n` +
            `[CONTEXT: PrimaryAudience=${aud.primaryAudience}, SecondaryAudience=${aud.secondaryAudience}, Ages=16-65, Attitudes=${aud.attitudes}, Context=${aud.context}, Format=${aud.format}, InformationNeed=${aud.informationNeed}, InformationMostImportant=${aud.informationMostImportant}, InformationMustUnderstand=${aud.informationMustUnderstand}, Purpose=${aud.purpose}, PurposeUse=${aud.purposeUse}, Urgency=${aud.urgency}]\n` +
            `Analyze this selection and output JSON only:\n${selection.text}`;
          directLine.postActivity({ from: { id: "user" }, type: "message", text: prompt }).subscribe({
            error: (sendErr: any) => {
              sub.unsubscribe();
              const sendMsg = sendErr?.message ? String(sendErr.message) : "Failed to send activity to bot.";
              reject(new Error(sendMsg));
            }
          });

          setTimeout(() => {
            sub.unsubscribe();
            reject(new Error("Timeout waiting for agent response."));
          }, 30000);
        });

        setRawResponse(JSON.stringify(data, null, 2));
        const validation = validatePlan(data);
        if (!validation.valid) {
          setStatus(`Invalid plan: ${validation.reason}`);
          setResults(null);
          return;
        }

        setResults(data);
        setStatus(`Plan ready. ${data.items.length} findings to apply.`);
      });
    } catch (err: any) {
      const message = err?.message ? String(err.message) : "Unknown error";
      setStatus(`Error: ${message}`);
    }
    setLoading(false);
  };

  const applyPlan = async () => {
    if (!results?.items?.length) {
      setStatus("No plan to apply.");
      return;
    }

    setLoading(true);
    setStatus("Applying highlights and notes...");
    setSkipped([]);
    setAppliedCount(0);

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        const localSkipped: { text: string; reason: string }[] = [];
        let applied = 0;

        for (const item of results.items) {
          const isSingleWord = item.match.text.trim().split(/\s+/).length === 1;
          const ranges = selection.search(item.match.text, { matchCase: false, matchWholeWord: isSingleWord });
          ranges.load("items");
          await context.sync();
          if (ranges.items.length === 0) {
            localSkipped.push({ text: item.match.text, reason: "Not found in selection" });
            continue;
          }

          const range = ranges.items[0];
          range.font.highlightColor = "Yellow";
          const label = item.note?.label || "Note";
          const message = item.note?.message || "";
          range.insertText(` [${label}, ${message}]`, "After");
          applied += 1;
        }

        await context.sync();
        setAppliedCount(applied);
        setSkipped(localSkipped);
        if (localSkipped.length > 0) {
          setStatus(`Applied ${applied}. Skipped ${localSkipped.length}.`);
        } else {
          setStatus(`Applied ${applied} highlights and notes.`);
        }
      });
    } catch (err) {
      setStatus("Error applying plan.");
    }

    setLoading(false);
  };

  return (
    <div className={styles.container}>
      {!isReady ? (
        /* STEP 1: INITIALIZATION FORM [cite: 42, 533] */
        <Card className={styles.setupCard}>
          <Text size={500} weight="bold" className={styles.headerText}>Audience Setup</Text>
          <Text size={200}>Required per CAN-ASC-3.1:2025 Section 10</Text>
          <Divider />

          <Field label="Who is the primary audience? " required>
            <Dropdown placeholder="Select audience..." onOptionSelect={(_, d) => setAud({...aud, primaryAudience: d.optionText!})}>
              <Option>Clients/Newcomers</Option>
              <Option>Internal Staff</Option>
              <Option>General Public</Option>
            </Dropdown>
          </Field>

          <Field label="Who is the secondary audience? " required>
            <Textarea 
              placeholder="e.g., Caregivers, family members, managers..." 
              value={aud.secondaryAudience}
              onChange={(_, d) => setAud({...aud, secondaryAudience: d.value})} 
            />
          </Field>

          <Field label="Urgency Level" required>
            <RadioGroup layout="horizontal" onChange={(_, d) => setAud({...aud, urgency: d.value})}>
              <Radio value="Low" label="Low" />
              <Radio value="High" label="High" />
            </RadioGroup>
          </Field>

          <Field label="Audience Attitudes/Concerns" required>
            <Textarea 
              placeholder="e.g., Stressed about services..." 
              value={aud.attitudes}
              onChange={(_, d) => setAud({...aud, attitudes: d.value})} 
            />
          </Field>

          <Field label="Context/Platform" required>
            <Dropdown placeholder="Where will they read this?" onOptionSelect={(_, d) => setAud({...aud, context: d.optionText!})}>
              <Option>Mobile Device</Option>
              <Option>Public Signage</Option>
              <Option>Official Document</Option>
            </Dropdown>
          </Field>

          <Field label="Format preference" required>
            <Dropdown placeholder="Preferred format" onOptionSelect={(_, d) => setAud({...aud, format: d.optionText!})}>
              <Option>Short paragraphs</Option>
              <Option>Bulleted list</Option>
              <Option>Step-by-step instructions</Option>
            </Dropdown>
          </Field>

          <Field label="Information the audience needs or wants" required>
            <Textarea 
              placeholder="What do they need to know?" 
              value={aud.informationNeed}
              onChange={(_, d) => setAud({...aud, informationNeed: d.value})} 
            />
          </Field>

          <Field label="Most important information" required>
            <Textarea 
              placeholder="What matters most to them?" 
              value={aud.informationMostImportant}
              onChange={(_, d) => setAud({...aud, informationMostImportant: d.value})} 
            />
          </Field>

          <Field label="Information they must understand" required>
            <Textarea 
              placeholder="What must be understood?" 
              value={aud.informationMustUnderstand}
              onChange={(_, d) => setAud({...aud, informationMustUnderstand: d.value})} 
            />
          </Field>

          <Field label="Primary Purpose" required>
            <Dropdown placeholder="What is the goal?" onOptionSelect={(_, d) => setAud({...aud, purpose: d.optionText!})}>
              <Option>Instructions</Option>
              <Option>Notification</Option>
              <Option>Legal Compliance</Option>
            </Dropdown>
          </Field>

          <Field label="Purpose and expected use" required>
            <Textarea 
              placeholder="How should they use the information?" 
              value={aud.purposeUse}
              onChange={(_, d) => setAud({...aud, purposeUse: d.value})} 
            />
          </Field>

          <Button 
            className={styles.primaryButton} 
            disabled={!isFormValid} // BLOCK START IF EMPTY [cite: 1537]
            onClick={() => { setIsReady(true); }}
          >
            Continue to Analysis
          </Button>
          {!isFormValid && <Text size={100} italic>Please complete all required fields.</Text>}
        </Card>
      ) : (
        /* STEP 2: ANALYSIS INTERFACE */
        <>
          <Card className={styles.setupCard}>
            <Button appearance="subtle" onClick={() => setIsReady(false)}>← Change Audience</Button>
            <Text block weight="semibold" style={{color: "#6200EE"}}>Target: {aud.primaryAudience}</Text>
            <Text block size={200}>Bot: {botStatus}</Text>
            <Button 
              className={styles.primaryButton} 
              disabled={loading || !directLine} 
              onClick={analyzeSelection}
            >
              {loading ? <Spinner size="tiny" label="Analyzing..." /> : (!directLine ? "Waiting for Bot Connection..." : "Analyze Selected Text")}
            </Button>
            <Button 
              className={styles.primaryButton}
              disabled={loading || !results?.items?.length}
              onClick={applyPlan}
            >
              Apply Plan
            </Button>
          </Card>

          <div style={{padding: "10px", backgroundColor: "#F3E5F5", borderRadius: "4px"}}>
            <Text weight="bold">Status: </Text><Text>{status}</Text>
            {results && <Text block weight="bold">Readability Score: {results.plainnessScore}%</Text>}
          </div>

          {results?.items.map((issue: any, i: number) => (
            <Card key={i} className={styles.issueItem}>
              <Text weight="bold" style={{color: "#9C27B0"}}>Issue #{i + 1}</Text>
              <Text block italic>"{issue.match.text}"</Text>
              <Text block size={200}>{issue.note.message}</Text>
              <Text block size={100}>Category: {issue.note.label || "Unlabeled"}</Text>
            </Card>
          ))}

          {skipped.length > 0 && (
            <Card className={styles.issueItem}>
              <Text weight="bold">Skipped Items</Text>
              {skipped.map((s, idx) => (
                <Text key={idx} block size={100}>
                  {s.text} - {s.reason}
                </Text>
              ))}
            </Card>
          )}

          {rawResponse && !results && (
            <Card className={styles.issueItem}>
              <Text weight="bold">Diagnostics (Raw Response)</Text>
              <Textarea readOnly rows={6} value={rawResponse} />
            </Card>
          )}
        </>
      )}
    </div>
  );
}