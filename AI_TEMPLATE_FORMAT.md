# AI Template Format Guide (Refined)

## The Problem (Old Format)

The old template used section headers like `ACTION ITEMS:` which AIs would sometimes treat as content, breaking the parsing. The new format uses **strict marker-based parsing** that's harder for AIs to misformat.

## New Format - @@ Markers

Each item must use a **category marker** (`@@ACTION`, `@@DECISION`, etc.) followed by pipe-delimited fields.
The full AI response should be one single line with multiple marker-items chained together.

### Category Markers

| Marker          | Type         | Fields                                  | Example                                                               |
| --------------- | ------------ | --------------------------------------- | --------------------------------------------------------------------- |
| `@@ACTION`      | Action Item  | room \| note \| assigned to \| due date | `@@ACTION: Electrical Room \| Replace fixtures \| John \| 2026-03-25` |
| `@@DECISION`    | Decision     | room \| note                            | `@@DECISION: HVAC Room \| Switched to variable frequency drives`      |
| `@@QUESTION`    | Question     | room \| note \| assigned to             | `@@QUESTION: Structural \| Load capacity? \| ""`                      |
| `@@OBSERVATION` | Observation  | room \| note                            | `@@OBSERVATION: Reception \| Wall thickness changed`                  |
| `@@UNASSIGNED`  | Unknown Room | note \| category                        | `@@UNASSIGNED: Follow up on permits \| QUESTION`                      |

## Critical Rules

1. **Entire output must be ONE single line** - Do NOT wrap to multiple lines
2. **Always use `@@` prefix** - Never skip the marker
3. **Pipe delimiter** - Use `|` to separate fields (ONLY between fields)
4. **Long text stays on same line** - If a note is long, keep it on that same line with spaces
5. **Empty fields okay** - Use `""` or just leave blank for missing values
6. **Room names must match** - Use UNASSIGNED if the room is uncertain
7. **No extra formatting** - No bullets, numbers, explanations, or section headers

## Valid Example Output (Single-Line Response)

```
@@ACTION: Electrical Room | Replace all south wall fixtures | Electrician | 2026-03-25 @@ACTION: HVAC Zone 3 | Pressurize test report | MEP Engineer | 2026-03-22 @@DECISION: Structural | Approved 8-inch reinforcement depth @@DECISION: Architectural | Finalize paint schedule next week @@QUESTION: Parking Level B | What's the new capacity requirement? | PM @@QUESTION: Safety | Fire rating on new sealant? | "" @@OBSERVATION: Lobby | Ceiling height increased by 2 feet @@OBSERVATION: Conference Room | Lighting fixtures are different from specs @@UNASSIGNED: Schedule progress meeting | ACTION @@UNASSIGNED: Document all RFIs | QUESTION
```

## Invalid Examples (Don't Do This)

❌ **Wrapping to multiple lines (THIS BREAKS PARSING RELIABILITY):**

```
@@ACTION: Electrical Room | Replace all south wall
fixtures that need attention | John | 2026-03-25
```

**FIX:** Keep the whole response on ONE line:

```
@@ACTION: Electrical Room | Replace all south wall fixtures that need attention | John | 2026-03-25 @@DECISION: HVAC | Switch to VFD
```

❌ **Missing markers:**

```
Electrical Room | Fix lights | John | 2026-03-25  ← Missing @@ACTION
```

❌ **Section headers included:**

```
ACTION ITEMS: Electrical Room | Fix lights | John | 2026-03-25
DECISIONS: HVAC | Switch to VFD  ← These break parsing
```

❌ **Wrong delimiter:**

```
@@ACTION - Electrical Room - Fix lights - John - 2026-03-25  ← Uses dashes instead of pipes
```

❌ **Multiple items per line:**

```
@@ACTION: Electrical Room | Fix lights | John | 2026-03-25; @@DECISION: HVAC | Switch to VFD  ← Two items on one line
```

## How to Use With AI

When you paste the template + transcript into your AI, provide:

1. **The template instructions** (which now have the `@@` marker format baked in)
2. **The project name and rooms**
3. **The meeting transcript or notes**

The new prompt now:

- Uses `@@` markers explicitly to mark each item
- Provides concrete examples
- Warns against common mistakes
- Asks the AI to output ONLY the data lines (no explanations)

## Testing the New Format

When you paste the template into your AI, you should get output like one single line:

```
@@ACTION: Master Bedroom | Install outlet covers | Electrician | 2026-03-27 @@DECISION: Living Room | Approve light fixture placement @@QUESTION: Kitchen | What's the outlet spacing code? | "" @@OBSERVATION: Basement | Drain tile needs extension
```

**NOT** like the old format with section headers between items:

```
ACTION ITEMS:
Master Bedroom | Install outlet covers | Electrician | 2026-03-27
Living Room | Light fixture placement | Electrician | 2026-03-29
```

## Critical Issue: Wrapped Text (Multi-Line Notes)

**The parser only captures ONE LINE per item.** If an AI wraps text to multiple lines, only the first line is imported.

❌ **This breaks importing (text gets cut off):**

```
@@ACTION: Electrical Room | Replace all south wall
fixtures that need attention | John | 2026-03-25
```

✅ **This works (everything on one line):**

```
@@ACTION: Electrical Room | Replace all south wall fixtures that need attention | John | 2026-03-25
```

The template now explicitly tells AIs: **"Each item MUST be on ONE SINGLE LINE - NEVER WRAP TO MULTIPLE LINES"** with warning symbols.

## If the AI Still Messes Up

If you find the AI is not following the format:

1. **Multi-line text issue** - Make sure AI outputs the entire response on ONE line
2. Try a different AI model/provider (some are more strict about instructions)
3. Ask the AI explicitly: "Keep every item on a single line, do not wrap text"
4. Ask the AI to output in "raw text mode" or "code block"
5. Test with a short transcript first before using long ones

---

**Ready to test?** Use the "Copy AI Template" button in the Redline tool and paste it into your AI with the meeting transcript. The output should match the format above exactly, with each item on a single line.
