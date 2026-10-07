| Persona | Agent sharing | Site | Site permission | Page | Agent A via Agents SDK | Agent B via Agents SDK | Agent B via hosted embed |
|---|---|---|---|---|---|---|---|
| User A | Agent A + Agent B | Site A | Visitor (Read) | OK | WORKS | NOT SUPPORTED | ANSWERS (no auth enforced) |
| User A | Agent A + Agent B | Site B | Member (Edit) | OK | WORKS | NOT SUPPORTED | ANSWERS (no auth enforced) |
| User A | Agent A + Agent B | Site C | Owner (Full control) | OK | WORKS | NOT SUPPORTED | ANSWERS (no auth enforced) |
| User B | Agent A only | Site A | Member (Edit) | OK | WORKS | NOT SUPPORTED | ANSWERS (no auth enforced) |
| User B | Agent A only | Site B | Visitor (Read) | OK | WORKS | NOT SUPPORTED | ANSWERS (no auth enforced) |
| User B | Agent A only | Site C | No access | ACCESS DENIED | - | - | - |
| User C | Agent B only | Site A | No access | ACCESS DENIED | - | - | - |
| User C | Agent B only | Site B | Member (Edit) | OK | BLOCKED - not shared | NOT SUPPORTED | ANSWERS (no auth enforced) |
| User C | Agent B only | Site C | Visitor (Read) | OK | BLOCKED - not shared | NOT SUPPORTED | ANSWERS (no auth enforced) |
| User D | No agent access | Site A | Visitor (Read) | OK | BLOCKED - not shared | NOT SUPPORTED | ANSWERS (no auth enforced) |
| User D | No agent access | Site B | No access | ACCESS DENIED | - | - | - |
| User D | No agent access | Site C | Member (Edit) | OK | BLOCKED - not shared | NOT SUPPORTED | ANSWERS (no auth enforced) |
| MarioR (maker/owner) | Owner of A + B | Site A | Owner | OK | WORKS | NOT SUPPORTED | ANSWERS (no auth enforced) |
| MarioR (maker/owner) | Owner of A + B | Site B | Owner | OK | WORKS | NOT SUPPORTED | ANSWERS (no auth enforced) |
| MarioR (maker/owner) | Owner of A + B | Site C | Owner | OK | WORKS | NOT SUPPORTED | ANSWERS (no auth enforced) |
