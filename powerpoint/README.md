# Grover search in PowerPoint

`grover-3-qubit-morph.pptx` steps through Grover's search for 3 qubits with
|101⟩ marked, one click per operation. The drawing (Grover plane and
probability bars) fills the left half of each 16:9 slide; the right half is
left blank for your own content.

| Slide | Step | k | P(marked) |
|---|---|---|---|
| 1 | Start, equal superposition \|+++⟩ | 0 | 12.5% |
| 2 | Oracle | 0.5 | 12.5% |
| 3 | Diffusion | 1 | 78.1% |
| 4 | Oracle | 1.5 | 78.1% |
| 5 | Diffusion | 2 | 94.5% (first peak) |
| 6 | Oracle | 2.5 | 94.5% |
| 7 | Diffusion | 3 | 33.0% (overshoot) |

Slides 2–7 use the **Morph** transition, so advancing rotates the state
vector, turns the mirror line and resizes the bars. Morph needs PowerPoint
2019, Microsoft 365 or PowerPoint for the web; older versions, Keynote and
Google Slides show a fade instead. Delete slides 6–7 to stop at the peak.

To copy the animation into another deck, copy the slides in order and keep
the Morph transition on each one. The moving shapes have names starting with
`!!` (for example `!!state-vector`); Morph uses those names to pair shapes
between slides, so don't rename them.

Each slide's speaker notes describe the step.

## Regenerating

```bash
pip install python-pptx
python3 build_grover_pptx.py
```
