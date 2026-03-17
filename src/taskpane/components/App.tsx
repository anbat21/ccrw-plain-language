/* global Word */
import * as React from "react";
import { 
  Button, Text, Field, Dropdown, Option, RadioGroup, Radio, Input,
  makeStyles, tokens, Spinner, Card, Textarea, Divider, Checkbox
} from "@fluentui/react-components";
import { DirectLine } from "botframework-directlinejs";
import { telemetry } from "../../services/telemetry";

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
  },
  selectedIssueItem: {
    borderLeft: "4px solid #0b1f37",
    padding: "10px",
    marginBottom: "10px",
    backgroundColor: "#E3F2FD",
    borderTop: "2px solid #0b1f37"
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
  const [isBotConnected, setIsBotConnected] = React.useState(false);
  const [selectedIssueIds, setSelectedIssueIds] = React.useState<Set<number>>(new Set());
  const [analysisScopeId, setAnalysisScopeId] = React.useState<number | null>(null);
  const directLineRef = React.useRef<any>(null);
  const connectionSubRef = React.useRef<any>(null);
  const activitySubRef = React.useRef<any>(null);
  const reconnectingRef = React.useRef(false);
  const reconnectAttemptsRef = React.useRef(0);

  const toggleIssueSelection = React.useCallback((issueIndex: number) => {
    setSelectedIssueIds((prev) => {
      const next = new Set(prev);
      if (next.has(issueIndex)) {
        next.delete(issueIndex);
      } else {
        next.add(issueIndex);
      }
      return next;
    });
  }, []);

  const connectBot = React.useCallback(async (isReconnect = false) => {
    try {
      // Call the internal API created in server.js to get a fresh short-lived token
      const res = await fetch('/api/get-token');
      if (!res.ok) {
        setIsBotConnected(false);
        setBotStatus(`Error: Token API failed (${res.status})`);
        setStatus("Bot connection failed. Check server and /api/get-token.");
        return false;
      }

      const { token } = await res.json();
      if (!token) {
        setIsBotConnected(false);
        setBotStatus("Error: No token received");
        setStatus("Bot connection failed. Token response was empty.");
        return false;
      }

      const dl = new DirectLine({ token: token, webSocket: false });

      if (connectionSubRef.current) {
        connectionSubRef.current.unsubscribe();
      }
      if (activitySubRef.current) {
        activitySubRef.current.unsubscribe();
      }
      if (directLineRef.current && directLineRef.current !== dl) {
        try {
          directLineRef.current.end();
        } catch {
          // Ignore close errors from previous stale connection
        }
      }

      directLineRef.current = dl;
      setDirectLine(dl);
      setIsBotConnected(false);
      setBotStatus(isReconnect ? "Bot reconnecting..." : "Connecting to bot...");
      setStatus(isReconnect ? "Refreshing token and reconnecting..." : "System Ready.");

      // DirectLine only begins connecting once activity$ or postActivity is subscribed.
      activitySubRef.current = dl.activity$.subscribe({
        next: () => {
          // No-op. This subscription exists to start and keep the DirectLine session alive.
        },
        error: () => {
          setIsBotConnected(false);
        }
      });

      connectionSubRef.current = dl.connectionStatus$.subscribe((connectionStatus: number) => {
        // 2=Online, 3=ExpiredToken, 4=FailedToConnect, 5=Ended
        if (connectionStatus === 2) {
          reconnectAttemptsRef.current = 0;
          setIsBotConnected(true);
          setBotStatus("Bot connected");
          setStatus(isReconnect ? "Bot reconnected with a fresh token." : "System Ready.");
          return;
        }

        if (connectionStatus === 3 || connectionStatus === 4 || connectionStatus === 5) {
          setIsBotConnected(false);
          if (reconnectingRef.current) {
            return;
          }

          reconnectingRef.current = true;
          reconnectAttemptsRef.current += 1;
          setBotStatus("Bot reconnecting...");
          setStatus("Connection interrupted. Refreshing token and reconnecting...");

          // Calculate exponential backoff: 1s, 2s, 4s, 8s, etc (max 30s)
          const delayMs = Math.min(1000 * Math.pow(2, reconnectAttemptsRef.current - 1), 30000);

          setTimeout(() => {
            void connectBot(true).finally(() => {
              reconnectingRef.current = false;
            });
          }, delayMs);
        }
      });

      return true;
    } catch (error) {
      console.error("Error initializing bot:", error);
      setIsBotConnected(false);
      setBotStatus("Error: Bot connection failed");
      setStatus("Bot connection failed. Check server and secret configuration.");
      return false;
    }
  }, []);

  // Initialize Application Insights telemetry
  React.useEffect(() => {
    const instrumentationKey = process.env.REACT_APP_INSTRUMENTATION_KEY;
    if (instrumentationKey) {
      telemetry.initialize({
        instrumentationKey,
        enableAutoRouteTracking: true,
        enableUnhandledPromiseRejectionTracking: true,
        autoTrackPageVisitTime: true,
      });
    } else {
      console.warn('[App] Application Insights Instrumentation Key not found in environment variables');
    }
  }, []);

  // Initialize DirectLine with secure token from server and auto-reconnect behavior
  React.useEffect(() => {
    void connectBot(false);

    return () => {
      if (connectionSubRef.current) {
        connectionSubRef.current.unsubscribe();
      }
      if (activitySubRef.current) {
        activitySubRef.current.unsubscribe();
      }
      if (directLineRef.current) {
        try {
          directLineRef.current.end();
        } catch {
          // Ignore close errors during unmount
        }
      }
    };
  }, [connectBot]);

  // Initialize issue selection when results change (issue #1 selected by default)
  React.useEffect(() => {
    if (results?.items?.length > 0) {
      setSelectedIssueIds(new Set([0]));
    } else {
      setSelectedIssueIds(new Set());
    }
  }, [results]);

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
        typeof item.replacementText === "string" && item.replacementText.trim().length > 0 &&
        item.style?.color &&
        typeof item.note?.label === "string" &&
        item.note?.message;
    });
    if (validItems.length !== data.items.length) {
      return { valid: false, reason: "One or more items are invalid." };
    }
    return { valid: true, reason: "" };
  };

  const dedupePlanItems = (items: any[]) => {
    const seen = new Set<string>();

    return items.filter((item: any) => {
      const key = [
        item.match?.text?.trim().toLowerCase(),
        item.replacementText?.trim().toLowerCase(),
        item.note?.label?.trim().toLowerCase(),
        item.note?.message?.trim().toLowerCase()
      ].join("::");

      if (seen.has(key)) {
        return false;
      }

      seen.add(key);
      return true;
    });
  };
  const analyzeSelection = async () => {
    if (!directLine || !isBotConnected) {
      setStatus("Bot not connected yet. Please wait...");
      return;
    }

    telemetry.trackEvent('AnalysisStarted', {
      audience: aud.primaryAudience,
      audienceType: aud.audienceType,
    });

    setLoading(true);
    setStatus("Processing...");
    setResults(null);
    setRawResponse("");
    setSkipped([]);
    setAppliedCount(0);
    
    try {
      await Word.run(async (context) => {
        if (analysisScopeId !== null) {
          const existingScope = context.document.contentControls.getByIdOrNullObject(analysisScopeId);
          existingScope.load("isNullObject");
          await context.sync();
          if (!existingScope.isNullObject) {
            existingScope.delete(false);
            await context.sync();
          }
        }

        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          setStatus("Please select text in Word before analyzing.");
          telemetry.trackEvent('AnalysisError', {
            reason: 'NoTextSelected',
          });
          setLoading(false);
          return;
        }

        const analysisScope = selection.insertContentControl();
        analysisScope.tag = "ccrw-analysis-scope";
        (analysisScope as any).appearance = "Hidden";
        analysisScope.load("id");
        await context.sync();
        setAnalysisScopeId(analysisScope.id);

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
            `- EVERY issue must include replacementText with the exact text that should replace match.text.\n` +
            `- replacementText must be plain replacement text only. Do not include labels, quotes around the whole answer, or explanations.\n` +
            `- EVERY note.message must explain why the replacement helps, not act as the source of truth for the replacement.\n` +
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
            `{"version":"1.0","scope":"selection","plainnessScore":0,"items":[{"id":"A1","type":"highlight","match":{"strategy":"exactText","text":"EXACT_PHRASE"},"replacementText":"REPLACEMENT_TEXT","style":{"color":"yellow"},"note":{"label":"sentence-length","message":"Split this into two shorter sentences so the action is easier to follow."}}],"summary":{"total":1,"categories":{"passive-voice":0,"sentence-length":0,"jargon":0,"context-clarity":0,"cognitive-load":0,"distress-tone":0,"navigation":0,"format":0,"digital-instructions":0}}}\n\n` +
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
          telemetry.trackEvent('AnalysisError', {
            reason: 'InvalidPlan',
            detail: validation.reason,
          });
          setResults(null);
          return;
        }

        const uniqueItems = dedupePlanItems(data.items);
        const normalizedData = {
          ...data,
          items: uniqueItems,
          summary: {
            ...data.summary,
            total: uniqueItems.length
          }
        };

        setResults(normalizedData);
        if (uniqueItems.length === 0) {
          analysisScope.delete(false);
          await context.sync();
          setAnalysisScopeId(null);
          setStatus("Analysis complete. No issues found – your text is plain language ready!");
          telemetry.trackEvent('AnalysisCompleted', {
            findingsCount: 0,
            plainessScore: normalizedData.plainnessScore,
          });
        } else {
          setStatus(`Analysis complete. ${uniqueItems.length} findings ready. Select issues to apply.`);
          telemetry.trackEvent('AnalysisCompleted', {
            findingsCount: uniqueItems.length,
            plainessScore: normalizedData.plainnessScore,
          });
        }
      });
    } catch (err: any) {
      const message = err?.message ? String(err.message) : "Unknown error";
      setStatus(`Error: ${message}`);
      telemetry.trackException(err, 'AnalysisError');
    }
    setLoading(false);
  };

  const applyPlan = async () => {
    if (!results?.items?.length) {
      setStatus("No plan to apply.");
      return;
    }

    if (selectedIssueIds.size === 0) {
      setStatus("Please select at least one issue to apply.");
      return;
    }

    telemetry.trackEvent('TipsApplyStarted', {
      selectedCount: String(selectedIssueIds.size),
      totalCount: String(results.items.length),
    });

    setLoading(true);
    setStatus("Applying replacements...");
    setSkipped([]);
    setAppliedCount(0);

    try {
      let applied = 0;
      let localSkipped: { text: string; reason: string }[] = [];
      const appliedIndices = new Set<number>();

      await Word.run(async (context) => {
        let searchScope: Word.Range = context.document.getSelection();

        if (analysisScopeId !== null) {
          const analysisScope = context.document.contentControls.getByIdOrNullObject(analysisScopeId);
          analysisScope.load("isNullObject");
          await context.sync();

          if (!analysisScope.isNullObject) {
            searchScope = analysisScope.getRange();
          }
        }

        for (let idx = 0; idx < results.items.length; idx++) {
          // Only process selected items
          if (!selectedIssueIds.has(idx)) {
            continue;
          }

          const item = results.items[idx];
          const isSingleWord = item.match.text.trim().split(/\s+/).length === 1;
          const ranges = searchScope.search(item.match.text, { matchCase: false, matchWholeWord: isSingleWord });
          ranges.load("items");
          await context.sync();
          if (ranges.items.length === 0) {
            localSkipped.push({ text: item.match.text, reason: "Not found in analyzed text" });
            continue;
          }

          const range = ranges.items[0];
          const replacementText = typeof item.replacementText === "string" ? item.replacementText.trim() : "";

          if (replacementText) {
            // Replace the text with the explicit replacement returned by the agent
            range.insertText(replacementText, "Replace");
            applied += 1;
            appliedIndices.add(idx);
          } else {
            localSkipped.push({ text: item.match.text, reason: "Missing replacement text from analysis result" });
          }
        }

        await context.sync();
      });

      setAppliedCount(applied);
      setSkipped(localSkipped);

      if (applied === 0) {
        setStatus("No selected issues could be applied. The selected text may already have changed. Re-run analysis to get fresh findings.");
        telemetry.trackEvent('TipsApplyCompleted', {
          appliedCount: 0,
          skippedCount: String(localSkipped.length),
        });
        return;
      }

      const remainingItems = results.items.filter((_: any, idx: number) => !appliedIndices.has(idx));

      if (remainingItems.length > 0) {
        setResults({
          ...results,
          items: remainingItems,
          summary: {
            ...results.summary,
            total: remainingItems.length
          }
        });
        setSelectedIssueIds(new Set([0]));
      } else {
        if (analysisScopeId !== null) {
          await Word.run(async (context) => {
            const analysisScope = context.document.contentControls.getByIdOrNullObject(analysisScopeId);
            analysisScope.load("isNullObject");
            await context.sync();
            if (!analysisScope.isNullObject) {
              analysisScope.delete(false);
              await context.sync();
            }
          });
        }
        setResults(null);
        setSelectedIssueIds(new Set());
        setAnalysisScopeId(null);
      }

      if (localSkipped.length > 0) {
        setStatus(
          remainingItems.length > 0
            ? `Applied ${applied} replacements. ${remainingItems.length} finding(s) remain. ${localSkipped.length} selected item(s) could not be applied.`
            : `Applied ${applied} replacements. Re-run analysis to review the updated text. ${localSkipped.length} selected item(s) could not be applied.`
        );
      } else {
        setStatus(
          remainingItems.length > 0
            ? `Applied ${applied} replacements. ${remainingItems.length} finding(s) remain.`
            : "All selected fixes were applied. Run Analyze again to confirm no issues remain."
        );
      }

      telemetry.trackEvent('TipsApplyCompleted', {
        appliedCount: String(applied),
        skippedCount: String(localSkipped.length),
        remainingCount: String(remainingItems.length),
      });
    } catch (err) {
      setStatus("Error applying plan.");
      telemetry.trackException(err as Error, 'TipsApplyError');
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
            onClick={() => {
              telemetry.trackEvent('AudienceSetupCompleted', {
                primaryAudience: aud.primaryAudience,
                audienceType: aud.audienceType,
                communicationType: aud.whatCreating,
              });
              setIsReady(true);
            }}
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
              disabled={loading || !directLine || !isBotConnected} 
              onClick={analyzeSelection}
            >
              {loading ? <Spinner size="tiny" label="Analyzing..." /> : (!directLine || !isBotConnected ? "Waiting for Bot Connection..." : "Analyze Selected Text")}
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

          {results?.items && results.items.length === 0 && (
            <Card className={styles.issueItem}>
              <Text weight="bold" style={{color: "#0b1f37"}}>No Issues Found</Text>
              <Text block size={200}>Your text meets plain language standards for your audience. Great job!</Text>
            </Card>
          )}

          {results?.items.map((issue: any, i: number) => (
            <Card 
              key={i} 
              className={selectedIssueIds.has(i) ? styles.selectedIssueItem : styles.issueItem}
              onClick={() => toggleIssueSelection(i)}
              style={{ cursor: "pointer" }}
            >
              <div style={{ display: "flex", gap: "10px", alignItems: "flex-start" }}>
                <Checkbox 
                  checked={selectedIssueIds.has(i)}
                  onClick={(event) => event.stopPropagation()}
                  onChange={() => toggleIssueSelection(i)}
                />
                <div style={{ flex: 1 }}>
                  <Text weight="bold" style={{color: "#0b1f37"}}>Tip #{i + 1}</Text>
                  <Text block italic>"{issue.match.text}"</Text>
                  <Text block size={200}>{issue.note.message}</Text>
                  <Text block size={100}>Category: {issue.note.label || "Unlabeled"}</Text>
                </div>
              </div>
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