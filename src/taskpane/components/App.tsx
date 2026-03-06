/* global Word */
import * as React from "react";
import { 
  Button, Text, Field, Dropdown, Option, RadioGroup, Radio, Input,
  makeStyles, tokens, Spinner, Card, Textarea, Divider
} from "@fluentui/react-components";
import { DirectLine } from "botframework-directlinejs";

// Professional Styling: Dark Navy and White Theme
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
    borderTop: "6px solid #0b1f37", 
    boxShadow: tokens.shadow16 
  },
  primaryButton: { 
    backgroundColor: "#0b1f37", 
    color: "white", 
    marginTop: "10px",
    ":hover": { backgroundColor: "#092033" },
    ":disabled": { backgroundColor: "#0b1f37", opacity: 0.55 }
  },
  headerText: { color: "#0b1f37", marginBottom: "5px" },
  helperText: { marginBottom: "10px", color: "#333" },
  fieldHintText: { display: "block", marginBottom: "6px", color: "#5f5f5f" },
  issueItem: { 
    borderLeft: "4px solid #0b1f37", 
    padding: "10px", 
    marginBottom: "10px", 
    backgroundColor: "white" 
  }
});

// Audience Setup State Type
type AudienceType = "" | "Internal" | "External";

type AudienceSetup = {
  topicMainMessage: string;
  audienceType: AudienceType;
  primaryAudience: string;
  secondaryAudience: string; // optional
  whereReadUse: string;
  whatCreating: string;
  timing: string;
  infoNeed: string;
  infoUnderstand: string;
  audienceDo: string;
};

// Conditional audience options by type
const audienceOptionsByType: Record<"Internal" | "External", string[]> = {
  Internal: [
    "All employees",
    "Frontline staff",
    "People leaders",
    "Executives",
    "HR and People and Culture",
    "IT and Digital",
    "Finance",
    "Operations",
    "Legal",
    "Communications and marketing",
    "Training, learning and development",
    "New hires",
    "Volunteers, Casual workers"
  ],
  External: [
    "Clients, service users",
    "Job seekers, candidates",
    "Employers, customers",
    "Partners, vendors, funders",
    "Public, community members",
    "Government, regulators"
  ]
};

// Helper function to build audience context for the AI engine
const buildAudienceContext = (aud: AudienceSetup) => {
  const lines: string[] = [];

  lines.push(`The topic and main message of this communication is, ${aud.topicMainMessage}.`);
  lines.push(`This communication is for an ${aud.audienceType} audience to the organization.`);
  lines.push(`The primary audience for this communication is ${aud.primaryAudience}.`);

  if (aud.secondaryAudience) {
    lines.push(`The secondary audience for this communication is ${aud.secondaryAudience}.`);
  } else {
    lines.push(`There is no secondary audience for this communication.`);
  }

  lines.push(`The information will be presented through a ${aud.whereReadUse}.`);
  lines.push(`This information is being presented in a ${aud.whatCreating}.`);
  lines.push(`The timing and context for this communication is ${aud.timing}.`);
  lines.push(`The audience is looking for ${aud.infoNeed}.`);
  lines.push(`The author wants the audience to understand, ${aud.infoUnderstand}.`);
  lines.push(`After receiving this information, the audience will need to ${aud.audienceDo}.`);

  return lines.join("\n");
};

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

  // Initialization Form Data
  const [aud, setAud] = React.useState<AudienceSetup>({
    topicMainMessage: "",
    audienceType: "",
    primaryAudience: "",
    secondaryAudience: "",
    whereReadUse: "",
    whatCreating: "",
    timing: "",
    infoNeed: "",
    infoUnderstand: "",
    audienceDo: ""
  });

  // Derived audience options based on selected type
  const primaryAudienceOptions =
    aud.audienceType ? audienceOptionsByType[aud.audienceType] : [];
  const secondaryAudienceOptions =
    aud.audienceType ? audienceOptionsByType[aud.audienceType] : [];

  // FORM VALIDATION: Secondary audience is optional
  const isFormValid =
    aud.topicMainMessage.trim().length > 0 &&
    !!aud.audienceType &&
    !!aud.primaryAudience &&
    !!aud.whereReadUse &&
    !!aud.whatCreating &&
    !!aud.timing &&
    aud.infoNeed.trim().length > 0 &&
    aud.infoUnderstand.trim().length > 0 &&
    !!aud.audienceDo;

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

        // Build audience context for the AI engine
        const audienceContext = buildAudienceContext(aud);

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

          const prompt = `You are the AI Analysis Engine for the CCRW Plain Language Reviewer. Return ONLY JSON.\n\n` +
            `Analyze the text for:\n` +
            `1) Passive voice (suggest active alternative)\n` +
            `2) Long sentences (>12 words) (suggest how to split)\n` +
            `3) Complex jargon (suggest simpler words)\n\n` +
            `Add on barrier checks, always run these even if the user does not disclose needs:\n` +
            `4) Contextless clarity (language and comprehension barriers)\n` +
            `   - If the sentence uses "this/that/it/they" or other unclear references, replace with a specific noun.\n` +
            `   - If the sentence uses an acronym or shorthand, define it in place the first time.\n\n` +
            `5) Cognitive load (memory, attention, and processing barriers)\n` +
            `   - If one sentence contains multiple required actions or stacked conditions, split into separate sentences, one action each.\n` +
            `   - If a deadline exists, keep it in the same sentence as the action it applies to.\n\n` +
            `6) Neutral tone (emotional and distress-related barriers)\n` +
            `   - If the sentence includes blame, shame, or threat language, rewrite as a neutral status plus next step.\n\n` +
            `7) Findability (information access and navigation barriers)\n` +
            `   - If the sentence points to "here/above/below/click here/more info" without a destination cue, rewrite with descriptive link text or a searchable label.\n\n` +
            `8) Format independence (visual, hearing, and format barriers)\n` +
            `   - If the sentence relies on visuals or audio, add a short text equivalent, and remove color or position-only directions.\n\n` +
            `9) Usable digital steps (digital and interactive accessibility barriers)\n` +
            `   - If the sentence gives vague portal or form instructions, name the field, button label, and expected input format.\n\n` +
            `Rules:\n` +
            `- EVERY issue must include the exact phrase from input text in match.text.\n` +
            `- EVERY note.message must contain a specific replacement suggestion, not just advice.\n` +
            `- Include a category label in note.label.\n` +
            `- Output must be STRICT JSON only.\n\n` +
            `Category labels to use in note.label:\n` +
            `- "passive-voice"\n` +
            `- "sentence-length"\n` +
            `- "jargon"\n` +
            `- "context-clarity"\n` +
            `- "cognitive-load"\n` +
            `- "distress-tone"\n` +
            `- "navigation"\n` +
            `- "format"\n` +
            `- "digital-instructions"\n\n` +
            `Return ONLY this JSON schema:\n` +
            `{"version":"1.0","scope":"selection","plainnessScore":0,"items":[{"id":"A1","type":"highlight","match":{"strategy":"exactText","text":"EXACT_PHRASE"},"style":{"color":"yellow"},"note":{"label":"sentence-length","message":"Replace with 'First sentence... Second sentence...' for clarity"}}],"summary":{"total":1,"categories":{"passive-voice":0,"sentence-length":0,"jargon":0,"context-clarity":0,"cognitive-load":0,"distress-tone":0,"navigation":0,"format":0,"digital-instructions":0}}}\n\n` +
            `[CONTEXT]\n${audienceContext}\n\n` +
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
    setStatus("Applying replacements...");
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
          const message = item.note?.message || "";
          
          // Parse suggestion from message (e.g., "Replace with 'xyz' for clarity")
          const suggestionMatch = message.match(/Replace with ['"](.+?)['"]|Replace with: ['"](.+?)['"]/i);
          const suggestion = suggestionMatch ? (suggestionMatch[1] || suggestionMatch[2]) : null;

          if (suggestion) {
            // Replace the text with suggestion
            range.insertText(suggestion, "Replace");
            applied += 1;
          } else {
            // Fallback: highlight + insert comment if no parseable suggestion
            range.font.highlightColor = "Yellow";
            const label = item.note?.label || "Note";
            range.insertText(` [${label}: ${message}]`, "After");
            applied += 1;
          }
        }

        await context.sync();
        setAppliedCount(applied);
        setSkipped(localSkipped);
        if (localSkipped.length > 0) {
          setStatus(`Applied ${applied} replacements. Skipped ${localSkipped.length}.`);
        } else {
          setStatus(`Applied ${applied} replacements successfully.`);
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
        /* STEP 1: AUDIENCE SETUP FORM */
        <Card className={styles.setupCard}>
          <Text size={500} weight="bold" className={styles.headerText}>
            Audience Setup
          </Text>

          <Text size={200} className={styles.helperText}>
            Text is only in plain language if it is tailored to its intended audience. Complete the form below to ensure that your recommendations are tailored for your intended audience.
          </Text>

          <Divider />

          <Field
            label="What is the topic and main message in one sentence?"
            required
          >
            <Text size={200} className={styles.fieldHintText}>
              Provide a brief introduction to the information you are communicating.
            </Text>
            <Input
              placeholder="Write one sentence that captures the topic and main message"
              value={aud.topicMainMessage}
              onChange={(_, d) => setAud((prev) => ({ ...prev, topicMainMessage: d.value }))}
            />
          </Field>

          <Field
            label="Audience Type"
            required
          >
            <Text size={200} className={styles.fieldHintText}>
              Indicate whether this communication is for an internal or external audience.
            </Text>
            <RadioGroup
              layout="horizontal"
              value={aud.audienceType}
              onChange={(_, d) => {
                const nextType = d.value as "Internal" | "External";
                setAud((prev) => ({
                  ...prev,
                  audienceType: nextType,
                  primaryAudience: "",
                  secondaryAudience: ""
                }));
              }}
            >
              <Radio value="Internal" label="Internal" />
              <Radio value="External" label="External" />
            </RadioGroup>
          </Field>

          <Field
            label="Who is the primary audience?"
            required
          >
            <Text size={200} className={styles.fieldHintText}>
              This is the primary consumer of the information you are communicating.
            </Text>
            <Dropdown
              placeholder={aud.audienceType ? "Select an audience..." : "Select an audience type first"}
              disabled={!aud.audienceType}
              value={aud.primaryAudience}
              selectedOptions={aud.primaryAudience ? [aud.primaryAudience] : []}
              onOptionSelect={(_, d) => setAud((prev) => ({ ...prev, primaryAudience: d.selectedOptions[0] ?? "" }))}
            >
              {primaryAudienceOptions.map((opt) => (
                <Option key={opt} value={opt}>{opt}</Option>
              ))}
            </Dropdown>
          </Field>

          <Field
            label="Who is the secondary audience?"
          >
            <Text size={200} className={styles.fieldHintText}>
              This is another audience who will be receiving this information, but they are not the primary group you are communicating with.
            </Text>
            <Dropdown
              placeholder={aud.audienceType ? "Select an audience, or leave blank" : "Select an audience type first"}
              disabled={!aud.audienceType}
              value={aud.secondaryAudience}
              selectedOptions={aud.secondaryAudience ? [aud.secondaryAudience] : []}
              onOptionSelect={(_, d) => setAud((prev) => ({ ...prev, secondaryAudience: d.selectedOptions[0] ?? "" }))}
            >
              <Option key="none" value="">
                No secondary audience
              </Option>

              {secondaryAudienceOptions.map((opt) => (
                <Option key={opt} value={opt}>{opt}</Option>
              ))}
            </Dropdown>
          </Field>

          <Field
            label="Where will people read or use this information?"
            required
          >
            <Text size={200} className={styles.fieldHintText}>
              This is how the communication will be presented.
            </Text>
            <Dropdown
              placeholder="Select one..."
              value={aud.whereReadUse}
              selectedOptions={aud.whereReadUse ? [aud.whereReadUse] : []}
              onOptionSelect={(_, d) => setAud((prev) => ({ ...prev, whereReadUse: d.selectedOptions[0] ?? "" }))}
            >
              <Option value="Web page, public website">Web page, public website</Option>
              <Option value="Portal, logged in website">Portal, logged in website</Option>
              <Option value="PDF attachment">PDF attachment</Option>
              <Option value="Printed handout, mailed letter">Printed handout, mailed letter</Option>
              <Option value="Poster, signage">Poster, signage</Option>
              <Option value="Mobile app screen">Mobile app screen</Option>
              <Option value="Chatbot response, virtual assistant response">Chatbot response, virtual assistant response</Option>
              <Option value="Call script, phone script">Call script, phone script</Option>
            </Dropdown>
          </Field>

          <Field
            label="What are you creating?"
            required
          >
            <Text size={200} className={styles.fieldHintText}>
              This is the type of communication you are developing.
            </Text>
            <Dropdown
              placeholder="Select one..."
              value={aud.whatCreating}
              selectedOptions={aud.whatCreating ? [aud.whatCreating] : []}
              onOptionSelect={(_, d) => setAud((prev) => ({ ...prev, whatCreating: d.selectedOptions[0] ?? "" }))}
            >
              <Option value="Email">Email</Option>
              <Option value="Intranet page">Intranet page</Option>
              <Option value="Chat">Chat</Option>
              <Option value="Newsletter">Newsletter</Option>
              <Option value="Policy manual, employee handbook">Policy manual, employee handbook</Option>
              <Option value="Procedure">Procedure</Option>
              <Option value="Form">Form</Option>
              <Option value="Memo">Memo</Option>
              <Option value="Report">Report</Option>
              <Option value="Job posting">Job posting</Option>
              <Option value="Training guide">Training guide</Option>
              <Option value="FAQ">FAQ</Option>
              <Option value="Customer letter">Customer letter</Option>
              <Option value="Press release">Press release</Option>
              <Option value="Meeting agenda">Meeting agenda</Option>
              <Option value="Slide notes">Slide notes</Option>
            </Dropdown>
          </Field>

          <Field
            label="Timing of Communication"
            required
          >
            <Text size={200} className={styles.fieldHintText}>
              This helps identify the context of when this information is being communicated to the audience.
            </Text>
            <Dropdown
              placeholder="Select one..."
              value={aud.timing}
              selectedOptions={aud.timing ? [aud.timing] : []}
              onOptionSelect={(_, d) => setAud((prev) => ({ ...prev, timing: d.selectedOptions[0] ?? "" }))}
            >
              <Option value="One time announcement">One time announcement</Option>
              <Option value="Ongoing reference content">Ongoing reference content</Option>
              <Option value="Time sensitive update">Time sensitive update</Option>
              <Option value="Compliance deadline, safety critical">Compliance deadline, safety critical</Option>
              <Option value="High emotion context, benefits, health, accommodation, discipline">High emotion context, benefits, health, accommodation, discipline</Option>
            </Dropdown>
          </Field>

          <Field
            label="What information does the audience need or want?"
            required
          >
            <Text size={200} className={styles.fieldHintText}>
              Describe what it is that your audience is looking for.
            </Text>
            <Textarea
              placeholder="The audience is looking for..."
              value={aud.infoNeed}
              onChange={(_, d) => setAud((prev) => ({ ...prev, infoNeed: d.value }))}
            />
          </Field>

          <Field
            label="What information do you want the audience to understand?"
            required
          >
            <Text size={200} className={styles.fieldHintText}>
              Describe what it is that you want your audience to know after consuming this information.
            </Text>
            <Textarea
              placeholder="I want the audience to understand..."
              value={aud.infoUnderstand}
              onChange={(_, d) => setAud((prev) => ({ ...prev, infoUnderstand: d.value }))}
            />
          </Field>

          <Field
            label="What does the audience need to do with this information?"
            required
          >
            <Text size={200} className={styles.fieldHintText}>
              Select the option that matches the next steps the audience has to take after receiving this information.
            </Text>
            <Dropdown
              placeholder="Select one..."
              value={aud.audienceDo}
              selectedOptions={aud.audienceDo ? [aud.audienceDo] : []}
              onOptionSelect={(_, d) => setAud((prev) => ({ ...prev, audienceDo: d.selectedOptions[0] ?? "" }))}
            >
              <Option value="Take an action">Take an action</Option>
              <Option value="Make a decision">Make a decision</Option>
              <Option value="Follow steps or a process">Follow steps or a process</Option>
              <Option value="Know it for reference later">Know it for reference later</Option>
              <Option value="Share it with someone else">Share it with someone else</Option>
              <Option value="Ask for help or request something">Ask for help or request something</Option>
            </Dropdown>
          </Field>

          <Button
            className={styles.primaryButton}
            disabled={!isFormValid}
            onClick={() => setIsReady(true)}
          >
            Start Analysis
          </Button>

          {!isFormValid && (
            <Text size={100} italic>
              Please complete all required fields.
            </Text>
          )}
        </Card>
      ) : (
        /* STEP 2: ANALYSIS INTERFACE */
        <>
          <Card className={styles.setupCard}>
            <Button appearance="subtle" onClick={() => setIsReady(false)}>← Change Audience</Button>
            <Text block weight="semibold" style={{color: "#0b1f37"}}>Target: {aud.primaryAudience}</Text>
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
              <Text weight="bold" style={{color: "#0b1f37"}}>Issue #{i + 1}</Text>
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