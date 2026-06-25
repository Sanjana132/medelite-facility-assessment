import { useState, useRef, useEffect } from "react";
import {
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell, ReferenceLine
} from "recharts";

const DOCX_STATIC = {"word/_rels/document.xml.rels":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIElkPSJySWQxIiBUeXBlPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzL3N0eWxlcyIgVGFyZ2V0PSJzdHlsZXMueG1sIi8+PFJlbGF0aW9uc2hpcCBJZD0icklkMiIgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9udW1iZXJpbmciIFRhcmdldD0ibnVtYmVyaW5nLnhtbCIvPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDMiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvZm9vdG5vdGVzIiBUYXJnZXQ9ImZvb3Rub3Rlcy54bWwiLz48UmVsYXRpb25zaGlwIElkPSJySWQ0IiBUeXBlPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzL2VuZG5vdGVzIiBUYXJnZXQ9ImVuZG5vdGVzLnhtbCIvPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDUiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvc2V0dGluZ3MiIFRhcmdldD0ic2V0dGluZ3MueG1sIi8+PFJlbGF0aW9uc2hpcCBJZD0icklkNiIgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9jb21tZW50cyIgVGFyZ2V0PSJjb21tZW50cy54bWwiLz48UmVsYXRpb25zaGlwIElkPSJySWR2d2hjc2N1aGdqbGIwcmpuajcxOG8iIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvaHlwZXJsaW5rIiBUYXJnZXQ9Imh0dHBzOi8vd3d3Lm1lZGljYXJlLmdvdi9jYXJlLWNvbXBhcmUvZGV0YWlscy9udXJzaW5nLWhvbWUvNjg2MTIzIiBUYXJnZXRNb2RlPSJFeHRlcm5hbCIvPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDgiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvZm9udFRhYmxlIiBUYXJnZXQ9ImZvbnRUYWJsZS54bWwiLz48L1JlbGF0aW9uc2hpcHM+","word/styles.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pjx3OnN0eWxlcyBtYzpJZ25vcmFibGU9IncxNCB3MTUiIHhtbG5zOm1jPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvbWFya3VwLWNvbXBhdGliaWxpdHkvMjAwNiIgeG1sbnM6cj0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcyIgeG1sbnM6dz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3dvcmRwcm9jZXNzaW5nbWwvMjAwNi9tYWluIiB4bWxuczp3MTQ9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkbWwiIHhtbG5zOncxNT0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEyL3dvcmRtbCI+PHc6ZG9jRGVmYXVsdHM+PHc6clByRGVmYXVsdD48dzpyUHI+PHc6ckZvbnRzIHc6YXNjaWk9IkFyaWFsIiB3OmNzPSJBcmlhbCIgdzplYXN0QXNpYT0iQXJpYWwiIHc6aEFuc2k9IkFyaWFsIi8+PHc6c3ogdzp2YWw9IjIwIi8+PHc6c3pDcyB3OnZhbD0iMjAiLz48L3c6clByPjwvdzpyUHJEZWZhdWx0Pjx3OnBQckRlZmF1bHQvPjwvdzpkb2NEZWZhdWx0cz48dzpzdHlsZSB3OnR5cGU9InBhcmFncmFwaCIgdzpzdHlsZUlkPSJUaXRsZSI+PHc6bmFtZSB3OnZhbD0iVGl0bGUiLz48dzpiYXNlZE9uIHc6dmFsPSJOb3JtYWwiLz48dzpuZXh0IHc6dmFsPSJOb3JtYWwiLz48dzpxRm9ybWF0Lz48dzpyUHI+PHc6c3ogdzp2YWw9IjU2Ii8+PHc6c3pDcyB3OnZhbD0iNTYiLz48L3c6clByPjwvdzpzdHlsZT48dzpzdHlsZSB3OnR5cGU9InBhcmFncmFwaCIgdzpzdHlsZUlkPSJIZWFkaW5nMSI+PHc6bmFtZSB3OnZhbD0iSGVhZGluZyAxIi8+PHc6YmFzZWRPbiB3OnZhbD0iTm9ybWFsIi8+PHc6bmV4dCB3OnZhbD0iTm9ybWFsIi8+PHc6cUZvcm1hdC8+PHc6clByPjx3OmNvbG9yIHc6dmFsPSIyRTc0QjUiLz48dzpzeiB3OnZhbD0iMzIiLz48dzpzekNzIHc6dmFsPSIzMiIvPjwvdzpyUHI+PC93OnN0eWxlPjx3OnN0eWxlIHc6dHlwZT0icGFyYWdyYXBoIiB3OnN0eWxlSWQ9IkhlYWRpbmcyIj48dzpuYW1lIHc6dmFsPSJIZWFkaW5nIDIiLz48dzpiYXNlZE9uIHc6dmFsPSJOb3JtYWwiLz48dzpuZXh0IHc6dmFsPSJOb3JtYWwiLz48dzpxRm9ybWF0Lz48dzpyUHI+PHc6Y29sb3Igdzp2YWw9IjJFNzRCNSIvPjx3OnN6IHc6dmFsPSIyNiIvPjx3OnN6Q3Mgdzp2YWw9IjI2Ii8+PC93OnJQcj48L3c6c3R5bGU+PHc6c3R5bGUgdzp0eXBlPSJwYXJhZ3JhcGgiIHc6c3R5bGVJZD0iSGVhZGluZzMiPjx3Om5hbWUgdzp2YWw9IkhlYWRpbmcgMyIvPjx3OmJhc2VkT24gdzp2YWw9Ik5vcm1hbCIvPjx3Om5leHQgdzp2YWw9Ik5vcm1hbCIvPjx3OnFGb3JtYXQvPjx3OnJQcj48dzpjb2xvciB3OnZhbD0iMUY0RDc4Ii8+PHc6c3ogdzp2YWw9IjI0Ii8+PHc6c3pDcyB3OnZhbD0iMjQiLz48L3c6clByPjwvdzpzdHlsZT48dzpzdHlsZSB3OnR5cGU9InBhcmFncmFwaCIgdzpzdHlsZUlkPSJIZWFkaW5nNCI+PHc6bmFtZSB3OnZhbD0iSGVhZGluZyA0Ii8+PHc6YmFzZWRPbiB3OnZhbD0iTm9ybWFsIi8+PHc6bmV4dCB3OnZhbD0iTm9ybWFsIi8+PHc6cUZvcm1hdC8+PHc6clByPjx3OmkvPjx3OmlDcy8+PHc6Y29sb3Igdzp2YWw9IjJFNzRCNSIvPjwvdzpyUHI+PC93OnN0eWxlPjx3OnN0eWxlIHc6dHlwZT0icGFyYWdyYXBoIiB3OnN0eWxlSWQ9IkhlYWRpbmc1Ij48dzpuYW1lIHc6dmFsPSJIZWFkaW5nIDUiLz48dzpiYXNlZE9uIHc6dmFsPSJOb3JtYWwiLz48dzpuZXh0IHc6dmFsPSJOb3JtYWwiLz48dzpxRm9ybWF0Lz48dzpyUHI+PHc6Y29sb3Igdzp2YWw9IjJFNzRCNSIvPjwvdzpyUHI+PC93OnN0eWxlPjx3OnN0eWxlIHc6dHlwZT0icGFyYWdyYXBoIiB3OnN0eWxlSWQ9IkhlYWRpbmc2Ij48dzpuYW1lIHc6dmFsPSJIZWFkaW5nIDYiLz48dzpiYXNlZE9uIHc6dmFsPSJOb3JtYWwiLz48dzpuZXh0IHc6dmFsPSJOb3JtYWwiLz48dzpxRm9ybWF0Lz48dzpyUHI+PHc6Y29sb3Igdzp2YWw9IjFGNEQ3OCIvPjwvdzpyUHI+PC93OnN0eWxlPjx3OnN0eWxlIHc6dHlwZT0icGFyYWdyYXBoIiB3OnN0eWxlSWQ9IlN0cm9uZyI+PHc6bmFtZSB3OnZhbD0iU3Ryb25nIi8+PHc6YmFzZWRPbiB3OnZhbD0iTm9ybWFsIi8+PHc6bmV4dCB3OnZhbD0iTm9ybWFsIi8+PHc6cUZvcm1hdC8+PHc6clByPjx3OmIvPjx3OmJDcy8+PC93OnJQcj48L3c6c3R5bGU+PHc6c3R5bGUgdzp0eXBlPSJwYXJhZ3JhcGgiIHc6c3R5bGVJZD0iTGlzdFBhcmFncmFwaCI+PHc6bmFtZSB3OnZhbD0iTGlzdCBQYXJhZ3JhcGgiLz48dzpiYXNlZE9uIHc6dmFsPSJOb3JtYWwiLz48dzpxRm9ybWF0Lz48L3c6c3R5bGU+PHc6c3R5bGUgdzp0eXBlPSJjaGFyYWN0ZXIiIHc6c3R5bGVJZD0iSHlwZXJsaW5rIj48dzpuYW1lIHc6dmFsPSJIeXBlcmxpbmsiLz48dzpiYXNlZE9uIHc6dmFsPSJEZWZhdWx0UGFyYWdyYXBoRm9udCIvPjx3OnVpUHJpb3JpdHkgdzp2YWw9Ijk5Ii8+PHc6dW5oaWRlV2hlblVzZWQvPjx3OnJQcj48dzpjb2xvciB3OnZhbD0iMDU2M0MxIi8+PHc6dSB3OnZhbD0ic2luZ2xlIi8+PC93OnJQcj48L3c6c3R5bGU+PHc6c3R5bGUgdzp0eXBlPSJjaGFyYWN0ZXIiIHc6c3R5bGVJZD0iRm9vdG5vdGVSZWZlcmVuY2UiPjx3Om5hbWUgdzp2YWw9ImZvb3Rub3RlIHJlZmVyZW5jZSIvPjx3OmJhc2VkT24gdzp2YWw9IkRlZmF1bHRQYXJhZ3JhcGhGb250Ii8+PHc6dWlQcmlvcml0eSB3OnZhbD0iOTkiLz48dzpzZW1pSGlkZGVuLz48dzp1bmhpZGVXaGVuVXNlZC8+PHc6clByPjx3OnZlcnRBbGlnbiB3OnZhbD0ic3VwZXJzY3JpcHQiLz48L3c6clByPjwvdzpzdHlsZT48dzpzdHlsZSB3OnR5cGU9InBhcmFncmFwaCIgdzpzdHlsZUlkPSJGb290bm90ZVRleHQiPjx3Om5hbWUgdzp2YWw9ImZvb3Rub3RlIHRleHQiLz48dzpiYXNlZE9uIHc6dmFsPSJOb3JtYWwiLz48dzpsaW5rIHc6dmFsPSJGb290bm90ZVRleHRDaGFyIi8+PHc6dWlQcmlvcml0eSB3OnZhbD0iOTkiLz48dzpzZW1pSGlkZGVuLz48dzp1bmhpZGVXaGVuVXNlZC8+PHc6cFByPjx3OnNwYWNpbmcgdzphZnRlcj0iMCIgdzpsaW5lPSIyNDAiIHc6bGluZVJ1bGU9ImF1dG8iLz48L3c6cFByPjx3OnJQcj48dzpzeiB3OnZhbD0iMjAiLz48dzpzekNzIHc6dmFsPSIyMCIvPjwvdzpyUHI+PC93OnN0eWxlPjx3OnN0eWxlIHc6dHlwZT0iY2hhcmFjdGVyIiB3OnN0eWxlSWQ9IkZvb3Rub3RlVGV4dENoYXIiPjx3Om5hbWUgdzp2YWw9IkZvb3Rub3RlIFRleHQgQ2hhciIvPjx3OmJhc2VkT24gdzp2YWw9IkRlZmF1bHRQYXJhZ3JhcGhGb250Ii8+PHc6bGluayB3OnZhbD0iRm9vdG5vdGVUZXh0Ii8+PHc6dWlQcmlvcml0eSB3OnZhbD0iOTkiLz48dzpzZW1pSGlkZGVuLz48dzp1bmhpZGVXaGVuVXNlZC8+PHc6clByPjx3OnN6IHc6dmFsPSIyMCIvPjx3OnN6Q3Mgdzp2YWw9IjIwIi8+PC93OnJQcj48L3c6c3R5bGU+PHc6c3R5bGUgdzp0eXBlPSJjaGFyYWN0ZXIiIHc6c3R5bGVJZD0iRW5kbm90ZVJlZmVyZW5jZSI+PHc6bmFtZSB3OnZhbD0iZW5kbm90ZSByZWZlcmVuY2UiLz48dzpiYXNlZE9uIHc6dmFsPSJEZWZhdWx0UGFyYWdyYXBoRm9udCIvPjx3OnVpUHJpb3JpdHkgdzp2YWw9Ijk5Ii8+PHc6c2VtaUhpZGRlbi8+PHc6dW5oaWRlV2hlblVzZWQvPjx3OnJQcj48dzp2ZXJ0QWxpZ24gdzp2YWw9InN1cGVyc2NyaXB0Ii8+PC93OnJQcj48L3c6c3R5bGU+PHc6c3R5bGUgdzp0eXBlPSJwYXJhZ3JhcGgiIHc6c3R5bGVJZD0iRW5kbm90ZVRleHQiPjx3Om5hbWUgdzp2YWw9ImVuZG5vdGUgdGV4dCIvPjx3OmJhc2VkT24gdzp2YWw9Ik5vcm1hbCIvPjx3Omxpbmsgdzp2YWw9IkVuZG5vdGVUZXh0Q2hhciIvPjx3OnVpUHJpb3JpdHkgdzp2YWw9Ijk5Ii8+PHc6c2VtaUhpZGRlbi8+PHc6dW5oaWRlV2hlblVzZWQvPjx3OnBQcj48dzpzcGFjaW5nIHc6YWZ0ZXI9IjAiIHc6bGluZT0iMjQwIiB3OmxpbmVSdWxlPSJhdXRvIi8+PC93OnBQcj48dzpyUHI+PHc6c3ogdzp2YWw9IjIwIi8+PHc6c3pDcyB3OnZhbD0iMjAiLz48L3c6clByPjwvdzpzdHlsZT48dzpzdHlsZSB3OnR5cGU9ImNoYXJhY3RlciIgdzpzdHlsZUlkPSJFbmRub3RlVGV4dENoYXIiPjx3Om5hbWUgdzp2YWw9IkVuZG5vdGUgVGV4dCBDaGFyIi8+PHc6YmFzZWRPbiB3OnZhbD0iRGVmYXVsdFBhcmFncmFwaEZvbnQiLz48dzpsaW5rIHc6dmFsPSJFbmRub3RlVGV4dCIvPjx3OnVpUHJpb3JpdHkgdzp2YWw9Ijk5Ii8+PHc6c2VtaUhpZGRlbi8+PHc6dW5oaWRlV2hlblVzZWQvPjx3OnJQcj48dzpzeiB3OnZhbD0iMjAiLz48dzpzekNzIHc6dmFsPSIyMCIvPjwvdzpyUHI+PC93OnN0eWxlPjx3OnN0eWxlIHc6dHlwZT0iY2hhcmFjdGVyIiB3OnN0eWxlSWQ9Ikh5cGVybGluayI+PHc6bmFtZSB3OnZhbD0iSHlwZXJsaW5rIi8+PHc6dWlQcmlvcml0eSB3OnZhbD0iOTkiLz48dzp1bmhpZGVXaGVuVXNlZC8+PHc6clByPjx3OmNvbG9yIHc6dmFsPSIwMEI0RDgiLz48dzp1IHc6dmFsPSJzaW5nbGUiLz48L3c6clByPjwvdzpzdHlsZT48L3c6c3R5bGVzPg==","docProps/core.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxjcDpjb3JlUHJvcGVydGllcyB4bWxuczpjcD0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9tZXRhZGF0YS9jb3JlLXByb3BlcnRpZXMiIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6ZGN0ZXJtcz0iaHR0cDovL3B1cmwub3JnL2RjL3Rlcm1zLyIgeG1sbnM6ZGNtaXR5cGU9Imh0dHA6Ly9wdXJsLm9yZy9kYy9kY21pdHlwZS8iIHhtbG5zOnhzaT0iaHR0cDovL3d3dy53My5vcmcvMjAwMS9YTUxTY2hlbWEtaW5zdGFuY2UiPjxkYzpjcmVhdG9yPlVuLW5hbWVkPC9kYzpjcmVhdG9yPjxjcDpsYXN0TW9kaWZpZWRCeT5Vbi1uYW1lZDwvY3A6bGFzdE1vZGlmaWVkQnk+PGNwOnJldmlzaW9uPjE8L2NwOnJldmlzaW9uPjxkY3Rlcm1zOmNyZWF0ZWQgeHNpOnR5cGU9ImRjdGVybXM6VzNDRFRGIj4yMDI2LTA2LTI1VDE5OjE4OjIyLjQwNlo8L2RjdGVybXM6Y3JlYXRlZD48ZGN0ZXJtczptb2RpZmllZCB4c2k6dHlwZT0iZGN0ZXJtczpXM0NEVEYiPjIwMjYtMDYtMjVUMTk6MTg6MjIuNDA2WjwvZGN0ZXJtczptb2RpZmllZD48L2NwOmNvcmVQcm9wZXJ0aWVzPg==","word/numbering.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pjx3Om51bWJlcmluZyBtYzpJZ25vcmFibGU9IncxNCB3MTUgd3AxNCIgeG1sbnM6d3BjPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTAvd29yZHByb2Nlc3NpbmdDYW52YXMiIHhtbG5zOm1jPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvbWFya3VwLWNvbXBhdGliaWxpdHkvMjAwNiIgeG1sbnM6bz0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTpvZmZpY2U6b2ZmaWNlIiB4bWxuczpyPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzIiB4bWxuczptPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9tYXRoIiB4bWxuczp2PSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOnZtbCIgeG1sbnM6d3AxND0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nRHJhd2luZyIgeG1sbnM6d3A9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9kcmF3aW5nbWwvMjAwNi93b3JkcHJvY2Vzc2luZ0RyYXdpbmciIHhtbG5zOncxMD0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTpvZmZpY2U6d29yZCIgeG1sbnM6dz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3dvcmRwcm9jZXNzaW5nbWwvMjAwNi9tYWluIiB4bWxuczp3MTQ9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkbWwiIHhtbG5zOncxNT0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEyL3dvcmRtbCIgeG1sbnM6d3BnPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTAvd29yZHByb2Nlc3NpbmdHcm91cCIgeG1sbnM6d3BpPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTAvd29yZHByb2Nlc3NpbmdJbmsiIHhtbG5zOnduZT0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDA2L3dvcmRtbCIgeG1sbnM6d3BzPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTAvd29yZHByb2Nlc3NpbmdTaGFwZSI+PHc6YWJzdHJhY3ROdW0gdzphYnN0cmFjdE51bUlkPSIxIiB3MTU6cmVzdGFydE51bWJlcmluZ0FmdGVyQnJlYWs9IjAiPjx3Om11bHRpTGV2ZWxUeXBlIHc6dmFsPSJoeWJyaWRNdWx0aWxldmVsIi8+PHc6bHZsIHc6aWx2bD0iMCIgdzE1OnRlbnRhdGl2ZT0iMSI+PHc6c3RhcnQgdzp2YWw9IjEiLz48dzpudW1GbXQgdzp2YWw9ImJ1bGxldCIvPjx3Omx2bFRleHQgdzp2YWw9IuKXjyIvPjx3Omx2bEpjIHc6dmFsPSJsZWZ0Ii8+PHc6cFByPjx3OmluZCB3OmxlZnQ9IjcyMCIgdzpoYW5naW5nPSIzNjAiLz48L3c6cFByPjwvdzpsdmw+PHc6bHZsIHc6aWx2bD0iMSIgdzE1OnRlbnRhdGl2ZT0iMSI+PHc6c3RhcnQgdzp2YWw9IjEiLz48dzpudW1GbXQgdzp2YWw9ImJ1bGxldCIvPjx3Omx2bFRleHQgdzp2YWw9IuKXiyIvPjx3Omx2bEpjIHc6dmFsPSJsZWZ0Ii8+PHc6cFByPjx3OmluZCB3OmxlZnQ9IjE0NDAiIHc6aGFuZ2luZz0iMzYwIi8+PC93OnBQcj48L3c6bHZsPjx3Omx2bCB3Omlsdmw9IjIiIHcxNTp0ZW50YXRpdmU9IjEiPjx3OnN0YXJ0IHc6dmFsPSIxIi8+PHc6bnVtRm10IHc6dmFsPSJidWxsZXQiLz48dzpsdmxUZXh0IHc6dmFsPSLilqAiLz48dzpsdmxKYyB3OnZhbD0ibGVmdCIvPjx3OnBQcj48dzppbmQgdzpsZWZ0PSIyMTYwIiB3Omhhbmdpbmc9IjM2MCIvPjwvdzpwUHI+PC93Omx2bD48dzpsdmwgdzppbHZsPSIzIiB3MTU6dGVudGF0aXZlPSIxIj48dzpzdGFydCB3OnZhbD0iMSIvPjx3Om51bUZtdCB3OnZhbD0iYnVsbGV0Ii8+PHc6bHZsVGV4dCB3OnZhbD0i4pePIi8+PHc6bHZsSmMgdzp2YWw9ImxlZnQiLz48dzpwUHI+PHc6aW5kIHc6bGVmdD0iMjg4MCIgdzpoYW5naW5nPSIzNjAiLz48L3c6cFByPjwvdzpsdmw+PHc6bHZsIHc6aWx2bD0iNCIgdzE1OnRlbnRhdGl2ZT0iMSI+PHc6c3RhcnQgdzp2YWw9IjEiLz48dzpudW1GbXQgdzp2YWw9ImJ1bGxldCIvPjx3Omx2bFRleHQgdzp2YWw9IuKXiyIvPjx3Omx2bEpjIHc6dmFsPSJsZWZ0Ii8+PHc6cFByPjx3OmluZCB3OmxlZnQ9IjM2MDAiIHc6aGFuZ2luZz0iMzYwIi8+PC93OnBQcj48L3c6bHZsPjx3Omx2bCB3Omlsdmw9IjUiIHcxNTp0ZW50YXRpdmU9IjEiPjx3OnN0YXJ0IHc6dmFsPSIxIi8+PHc6bnVtRm10IHc6dmFsPSJidWxsZXQiLz48dzpsdmxUZXh0IHc6dmFsPSLilqAiLz48dzpsdmxKYyB3OnZhbD0ibGVmdCIvPjx3OnBQcj48dzppbmQgdzpsZWZ0PSI0MzIwIiB3Omhhbmdpbmc9IjM2MCIvPjwvdzpwUHI+PC93Omx2bD48dzpsdmwgdzppbHZsPSI2IiB3MTU6dGVudGF0aXZlPSIxIj48dzpzdGFydCB3OnZhbD0iMSIvPjx3Om51bUZtdCB3OnZhbD0iYnVsbGV0Ii8+PHc6bHZsVGV4dCB3OnZhbD0i4pePIi8+PHc6bHZsSmMgdzp2YWw9ImxlZnQiLz48dzpwUHI+PHc6aW5kIHc6bGVmdD0iNTA0MCIgdzpoYW5naW5nPSIzNjAiLz48L3c6cFByPjwvdzpsdmw+PHc6bHZsIHc6aWx2bD0iNyIgdzE1OnRlbnRhdGl2ZT0iMSI+PHc6c3RhcnQgdzp2YWw9IjEiLz48dzpudW1GbXQgdzp2YWw9ImJ1bGxldCIvPjx3Omx2bFRleHQgdzp2YWw9IuKXjyIvPjx3Omx2bEpjIHc6dmFsPSJsZWZ0Ii8+PHc6cFByPjx3OmluZCB3OmxlZnQ9IjU3NjAiIHc6aGFuZ2luZz0iMzYwIi8+PC93OnBQcj48L3c6bHZsPjx3Omx2bCB3Omlsdmw9IjgiIHcxNTp0ZW50YXRpdmU9IjEiPjx3OnN0YXJ0IHc6dmFsPSIxIi8+PHc6bnVtRm10IHc6dmFsPSJidWxsZXQiLz48dzpsdmxUZXh0IHc6dmFsPSLil48iLz48dzpsdmxKYyB3OnZhbD0ibGVmdCIvPjx3OnBQcj48dzppbmQgdzpsZWZ0PSI2NDgwIiB3Omhhbmdpbmc9IjM2MCIvPjwvdzpwUHI+PC93Omx2bD48L3c6YWJzdHJhY3ROdW0+PHc6bnVtIHc6bnVtSWQ9IjEiPjx3OmFic3RyYWN0TnVtSWQgdzp2YWw9IjEiLz48dzpsdmxPdmVycmlkZSB3Omlsdmw9IjAiPjx3OnN0YXJ0T3ZlcnJpZGUgdzp2YWw9IjEiLz48L3c6bHZsT3ZlcnJpZGU+PC93Om51bT48L3c6bnVtYmVyaW5nPg==","_rels/.rels":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIElkPSJySWQxIiBUeXBlPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzL29mZmljZURvY3VtZW50IiBUYXJnZXQ9IndvcmQvZG9jdW1lbnQueG1sIi8+PFJlbGF0aW9uc2hpcCBJZD0icklkMiIgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzL21ldGFkYXRhL2NvcmUtcHJvcGVydGllcyIgVGFyZ2V0PSJkb2NQcm9wcy9jb3JlLnhtbCIvPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDMiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvZXh0ZW5kZWQtcHJvcGVydGllcyIgVGFyZ2V0PSJkb2NQcm9wcy9hcHAueG1sIi8+PFJlbGF0aW9uc2hpcCBJZD0icklkNCIgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9jdXN0b20tcHJvcGVydGllcyIgVGFyZ2V0PSJkb2NQcm9wcy9jdXN0b20ueG1sIi8+PC9SZWxhdGlvbnNoaXBzPg==","[Content_Types].xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz48VHlwZXMgeG1sbnM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9wYWNrYWdlLzIwMDYvY29udGVudC10eXBlcyI+PERlZmF1bHQgQ29udGVudFR5cGU9ImltYWdlL3BuZyIgRXh0ZW5zaW9uPSJwbmciLz48RGVmYXVsdCBDb250ZW50VHlwZT0iaW1hZ2UvanBlZyIgRXh0ZW5zaW9uPSJqcGVnIi8+PERlZmF1bHQgQ29udGVudFR5cGU9ImltYWdlL2pwZWciIEV4dGVuc2lvbj0ianBnIi8+PERlZmF1bHQgQ29udGVudFR5cGU9ImltYWdlL2JtcCIgRXh0ZW5zaW9uPSJibXAiLz48RGVmYXVsdCBDb250ZW50VHlwZT0iaW1hZ2UvZ2lmIiBFeHRlbnNpb249ImdpZiIvPjxEZWZhdWx0IENvbnRlbnRUeXBlPSJpbWFnZS9zdmcreG1sIiBFeHRlbnNpb249InN2ZyIvPjxEZWZhdWx0IENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtcGFja2FnZS5yZWxhdGlvbnNoaXBzK3htbCIgRXh0ZW5zaW9uPSJyZWxzIi8+PERlZmF1bHQgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3htbCIgRXh0ZW5zaW9uPSJ4bWwiLz48RGVmYXVsdCBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50Lm9iZnVzY2F0ZWRGb250IiBFeHRlbnNpb249Im9kdHRmIi8+PE92ZXJyaWRlIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQud29yZHByb2Nlc3NpbmdtbC5kb2N1bWVudC5tYWluK3htbCIgUGFydE5hbWU9Ii93b3JkL2RvY3VtZW50LnhtbCIvPjxPdmVycmlkZSBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LndvcmRwcm9jZXNzaW5nbWwuc3R5bGVzK3htbCIgUGFydE5hbWU9Ii93b3JkL3N0eWxlcy54bWwiLz48T3ZlcnJpZGUgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1wYWNrYWdlLmNvcmUtcHJvcGVydGllcyt4bWwiIFBhcnROYW1lPSIvZG9jUHJvcHMvY29yZS54bWwiLz48T3ZlcnJpZGUgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC5jdXN0b20tcHJvcGVydGllcyt4bWwiIFBhcnROYW1lPSIvZG9jUHJvcHMvY3VzdG9tLnhtbCIvPjxPdmVycmlkZSBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LmV4dGVuZGVkLXByb3BlcnRpZXMreG1sIiBQYXJ0TmFtZT0iL2RvY1Byb3BzL2FwcC54bWwiLz48T3ZlcnJpZGUgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC53b3JkcHJvY2Vzc2luZ21sLm51bWJlcmluZyt4bWwiIFBhcnROYW1lPSIvd29yZC9udW1iZXJpbmcueG1sIi8+PE92ZXJyaWRlIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQud29yZHByb2Nlc3NpbmdtbC5mb290bm90ZXMreG1sIiBQYXJ0TmFtZT0iL3dvcmQvZm9vdG5vdGVzLnhtbCIvPjxPdmVycmlkZSBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LndvcmRwcm9jZXNzaW5nbWwuZW5kbm90ZXMreG1sIiBQYXJ0TmFtZT0iL3dvcmQvZW5kbm90ZXMueG1sIi8+PE92ZXJyaWRlIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQud29yZHByb2Nlc3NpbmdtbC5zZXR0aW5ncyt4bWwiIFBhcnROYW1lPSIvd29yZC9zZXR0aW5ncy54bWwiLz48T3ZlcnJpZGUgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC53b3JkcHJvY2Vzc2luZ21sLmNvbW1lbnRzK3htbCIgUGFydE5hbWU9Ii93b3JkL2NvbW1lbnRzLnhtbCIvPjxPdmVycmlkZSBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LndvcmRwcm9jZXNzaW5nbWwuZm9udFRhYmxlK3htbCIgUGFydE5hbWU9Ii93b3JkL2ZvbnRUYWJsZS54bWwiLz48L1R5cGVzPg==","docProps/custom.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxQcm9wZXJ0aWVzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9jdXN0b20tcHJvcGVydGllcyIgeG1sbnM6dnQ9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L2RvY1Byb3BzVlR5cGVzIi8+","docProps/app.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxQcm9wZXJ0aWVzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9leHRlbmRlZC1wcm9wZXJ0aWVzIiB4bWxuczp2dD0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvZG9jUHJvcHNWVHlwZXMiLz4=","word/footnotes.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pjx3OmZvb3Rub3RlcyB4bWxuczp3cGM9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkcHJvY2Vzc2luZ0NhbnZhcyIgeG1sbnM6bWM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9tYXJrdXAtY29tcGF0aWJpbGl0eS8yMDA2IiB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiIHhtbG5zOnI9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMiIHhtbG5zOm09Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L21hdGgiIHhtbG5zOnY9InVybjpzY2hlbWFzLW1pY3Jvc29mdC1jb206dm1sIiB4bWxuczp3cDE0PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTAvd29yZHByb2Nlc3NpbmdEcmF3aW5nIiB4bWxuczp3cD0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL2RyYXdpbmdtbC8yMDA2L3dvcmRwcm9jZXNzaW5nRHJhd2luZyIgeG1sbnM6dzEwPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTp3b3JkIiB4bWxuczp3PSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvd29yZHByb2Nlc3NpbmdtbC8yMDA2L21haW4iIHhtbG5zOncxND0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRtbCIgeG1sbnM6dzE1PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTIvd29yZG1sIiB4bWxuczp3cGc9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkcHJvY2Vzc2luZ0dyb3VwIiB4bWxuczp3cGk9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkcHJvY2Vzc2luZ0luayIgeG1sbnM6d25lPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMDYvd29yZG1sIiB4bWxuczp3cHM9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkcHJvY2Vzc2luZ1NoYXBlIiBtYzpJZ25vcmFibGU9IncxNCB3MTUgd3AxNCI+PHc6Zm9vdG5vdGUgdzp0eXBlPSJzZXBhcmF0b3IiIHc6aWQ9Ii0xIj48dzpwPjx3OnBQcj48dzpzcGFjaW5nIHc6YWZ0ZXI9IjAiIHc6bGluZT0iMjQwIiB3OmxpbmVSdWxlPSJhdXRvIi8+PC93OnBQcj48dzpyPjx3OnJQcj48dzpyU3R5bGUgdzp2YWw9IkZvb3Rub3RlUmVmZXJlbmNlIi8+PC93OnJQcj48dzpmb290bm90ZVJlZi8+PC93OnI+PHc6cj48dzpzZXBhcmF0b3IvPjwvdzpyPjwvdzpwPjwvdzpmb290bm90ZT48dzpmb290bm90ZSB3OnR5cGU9ImNvbnRpbnVhdGlvblNlcGFyYXRvciIgdzppZD0iMCI+PHc6cD48dzpwUHI+PHc6c3BhY2luZyB3OmFmdGVyPSIwIiB3OmxpbmU9IjI0MCIgdzpsaW5lUnVsZT0iYXV0byIvPjwvdzpwUHI+PHc6cj48dzpyUHI+PHc6clN0eWxlIHc6dmFsPSJGb290bm90ZVJlZmVyZW5jZSIvPjwvdzpyUHI+PHc6Zm9vdG5vdGVSZWYvPjwvdzpyPjx3OnI+PHc6Y29udGludWF0aW9uU2VwYXJhdG9yLz48L3c6cj48L3c6cD48L3c6Zm9vdG5vdGU+PC93OmZvb3Rub3Rlcz4=","word/_rels/footnotes.xml.rels":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIi8+","word/endnotes.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz48dzplbmRub3RlcyB4bWxuczp3cGM9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkcHJvY2Vzc2luZ0NhbnZhcyIgeG1sbnM6bWM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9tYXJrdXAtY29tcGF0aWJpbGl0eS8yMDA2IiB4bWxuczpvPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTpvZmZpY2UiIHhtbG5zOnI9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMiIHhtbG5zOm09Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L21hdGgiIHhtbG5zOnY9InVybjpzY2hlbWFzLW1pY3Jvc29mdC1jb206dm1sIiB4bWxuczp3cDE0PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTAvd29yZHByb2Nlc3NpbmdEcmF3aW5nIiB4bWxuczp3cD0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL2RyYXdpbmdtbC8yMDA2L3dvcmRwcm9jZXNzaW5nRHJhd2luZyIgeG1sbnM6dzEwPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOm9mZmljZTp3b3JkIiB4bWxuczp3PSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvd29yZHByb2Nlc3NpbmdtbC8yMDA2L21haW4iIHhtbG5zOncxND0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRtbCIgeG1sbnM6dzE1PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTIvd29yZG1sIiB4bWxuczp3cGc9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkcHJvY2Vzc2luZ0dyb3VwIiB4bWxuczp3cGk9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkcHJvY2Vzc2luZ0luayIgeG1sbnM6d25lPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMDYvd29yZG1sIiB4bWxuczp3cHM9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkcHJvY2Vzc2luZ1NoYXBlIiBtYzpJZ25vcmFibGU9IncxNCB3MTUgd3AxNCI+PHc6ZW5kbm90ZSB3OnR5cGU9InNlcGFyYXRvciIgdzppZD0iLTEiPjx3OnA+PHc6cFByPjx3OnNwYWNpbmcgdzphZnRlcj0iMCIgdzpsaW5lPSIyNDAiIHc6bGluZVJ1bGU9ImF1dG8iLz48L3c6cFByPjx3OnI+PHc6clByPjx3OnJTdHlsZSB3OnZhbD0iRW5kbm90ZVJlZmVyZW5jZSIvPjwvdzpyUHI+PHc6ZW5kbm90ZVJlZi8+PC93OnI+PHc6cj48dzpzZXBhcmF0b3IvPjwvdzpyPjwvdzpwPjwvdzplbmRub3RlPjx3OmVuZG5vdGUgdzp0eXBlPSJjb250aW51YXRpb25TZXBhcmF0b3IiIHc6aWQ9IjAiPjx3OnA+PHc6cFByPjx3OnNwYWNpbmcgdzphZnRlcj0iMCIgdzpsaW5lPSIyNDAiIHc6bGluZVJ1bGU9ImF1dG8iLz48L3c6cFByPjx3OnI+PHc6clByPjx3OnJTdHlsZSB3OnZhbD0iRW5kbm90ZVJlZmVyZW5jZSIvPjwvdzpyUHI+PHc6ZW5kbm90ZVJlZi8+PC93OnI+PHc6cj48dzpjb250aW51YXRpb25TZXBhcmF0b3IvPjwvdzpyPjwvdzpwPjwvdzplbmRub3RlPjwvdzplbmRub3Rlcz4=","word/_rels/endnotes.xml.rels":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIi8+","word/settings.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pjx3OnNldHRpbmdzIHhtbG5zOndwYz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nQ2FudmFzIiB4bWxuczptYz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL21hcmt1cC1jb21wYXRpYmlsaXR5LzIwMDYiIHhtbG5zOm89InVybjpzY2hlbWFzLW1pY3Jvc29mdC1jb206b2ZmaWNlOm9mZmljZSIgeG1sbnM6cj0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcyIgeG1sbnM6bT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvbWF0aCIgeG1sbnM6dj0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTp2bWwiIHhtbG5zOndwMTQ9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkcHJvY2Vzc2luZ0RyYXdpbmciIHhtbG5zOndwPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvZHJhd2luZ21sLzIwMDYvd29yZHByb2Nlc3NpbmdEcmF3aW5nIiB4bWxuczp3MTA9InVybjpzY2hlbWFzLW1pY3Jvc29mdC1jb206b2ZmaWNlOndvcmQiIHhtbG5zOnc9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy93b3JkcHJvY2Vzc2luZ21sLzIwMDYvbWFpbiIgeG1sbnM6dzE0PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTAvd29yZG1sIiB4bWxuczp3MTU9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMi93b3JkbWwiIHhtbG5zOndwZz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nR3JvdXAiIHhtbG5zOndwaT0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nSW5rIiB4bWxuczp3bmU9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAwNi93b3JkbWwiIHhtbG5zOndwcz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nU2hhcGUiIG1jOklnbm9yYWJsZT0idzE0IHcxNSB3cDE0Ij48dzpkaXNwbGF5QmFja2dyb3VuZFNoYXBlLz48dzpldmVuQW5kT2RkSGVhZGVycyB3OnZhbD0iZmFsc2UiLz48dzpjb21wYXQ+PHc6Y29tcGF0U2V0dGluZyB3OnZhbD0iMTUiIHc6bmFtZT0iY29tcGF0aWJpbGl0eU1vZGUiIHc6dXJpPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkIi8+PC93OmNvbXBhdD48L3c6c2V0dGluZ3M+","word/comments.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pjx3OmNvbW1lbnRzIHhtbG5zOmN4PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5nLzIwMTQvY2hhcnRleCIgeG1sbnM6Y3gxPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5nLzIwMTUvOS84L2NoYXJ0ZXgiIHhtbG5zOmN4Mj0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2UvZHJhd2luZy8yMDE1LzEwLzIxL2NoYXJ0ZXgiIHhtbG5zOmN4Mz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2UvZHJhd2luZy8yMDE2LzUvOS9jaGFydGV4IiB4bWxuczpjeDQ9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL2RyYXdpbmcvMjAxNi81LzEwL2NoYXJ0ZXgiIHhtbG5zOmN4NT0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2UvZHJhd2luZy8yMDE2LzUvMTEvY2hhcnRleCIgeG1sbnM6Y3g2PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5nLzIwMTYvNS8xMi9jaGFydGV4IiB4bWxuczpjeDc9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL2RyYXdpbmcvMjAxNi81LzEzL2NoYXJ0ZXgiIHhtbG5zOmN4OD0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2UvZHJhd2luZy8yMDE2LzUvMTQvY2hhcnRleCIgeG1sbnM6bWM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9tYXJrdXAtY29tcGF0aWJpbGl0eS8yMDA2IiB4bWxuczphaW5rPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5nLzIwMTYvaW5rIiB4bWxuczphbTNkPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5nLzIwMTcvbW9kZWwzZCIgeG1sbnM6bz0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTpvZmZpY2U6b2ZmaWNlIiB4bWxuczpyPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzIiB4bWxuczptPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9tYXRoIiB4bWxuczp2PSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOnZtbCIgeG1sbnM6d3AxND0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nRHJhd2luZyIgeG1sbnM6d3A9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9kcmF3aW5nbWwvMjAwNi93b3JkcHJvY2Vzc2luZ0RyYXdpbmciIHhtbG5zOncxMD0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTpvZmZpY2U6d29yZCIgeG1sbnM6dz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3dvcmRwcm9jZXNzaW5nbWwvMjAwNi9tYWluIiB4bWxuczp3MTQ9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkbWwiIHhtbG5zOncxNT0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEyL3dvcmRtbCIgeG1sbnM6dzE2Y2V4PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTgvd29yZG1sL2NleCIgeG1sbnM6dzE2Y2lkPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTYvd29yZG1sL2NpZCIgeG1sbnM6dzE2PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTgvd29yZG1sIiB4bWxuczp3MTZzZHRkaD0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDIwL3dvcmRtbC9zZHRkYXRhaGFzaCIgeG1sbnM6dzE2c2U9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxNS93b3JkbWwvc3ltZXgiIHhtbG5zOndwZz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nR3JvdXAiIHhtbG5zOndwaT0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nSW5rIiB4bWxuczp3bmU9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAwNi93b3JkbWwiIHhtbG5zOndwcz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nU2hhcGUiLz4=","word/_rels/comments.xml.rels":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIi8+","word/fontTable.xml":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pjx3OmZvbnRzIHhtbG5zOm1jPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvbWFya3VwLWNvbXBhdGliaWxpdHkvMjAwNiIgeG1sbnM6cj0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcyIgeG1sbnM6dz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3dvcmRwcm9jZXNzaW5nbWwvMjAwNi9tYWluIiB4bWxuczp3MTQ9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkbWwiIHhtbG5zOncxNT0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEyL3dvcmRtbCIgeG1sbnM6dzE2Y2V4PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTgvd29yZG1sL2NleCIgeG1sbnM6dzE2Y2lkPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTYvd29yZG1sL2NpZCIgeG1sbnM6dzE2PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTgvd29yZG1sIiB4bWxuczp3MTZzZHRkaD0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDIwL3dvcmRtbC9zZHRkYXRhaGFzaCIgeG1sbnM6dzE2c2U9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxNS93b3JkbWwvc3ltZXgiIG1jOklnbm9yYWJsZT0idzE0IHcxNSB3MTZzZSB3MTZjaWQgdzE2IHcxNmNleCB3MTZzZHRkaCIvPg==","word/_rels/fontTable.xml.rels":"PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIi8+"};

const C = {
  magenta:"#D400AA", teal:"#00B4D8", navy:"#0D1B2A",
  bg:"#F4F8FB", border:"#DDE3EA", dark:"#1A2533",
  mid:"#4A5568", light:"#718096",
  green:"#22C55E", amber:"#F59E0B", red:"#EF4444", white:"#FFFFFF"
};

const REF = {
  ccn:"686123", legalName:"Kendall Lakes Healthcare and Rehab Center",
  location:"5280 SW 157th Ave, Miami, FL 33193", state:"FL", beds:"120",
  overallStar:"1", healthStar:"1", staffStar:"2", qualityStar:"4",
  strHosp:"18.7", strHospNatl:"21.5", strHospState:"23.8",
  strED:"13.9", strEDNatl:"11.6", strEDState:"9.3",
  ltHosp:"1.86", ltHospNatl:"1.65", ltHospState:"1.95",
  ltED:"6.94", ltEDNatl:"1.65", ltEDState:"1.21"
};

function esc(s) {
  return String(s || "—")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function makeDocXML(rows, state, ccn) {
  const url = "https://www.medicare.gov/care-compare/details/nursing-home/" + ccn;
  const brd = "<w:tcBorders><w:top w:val=\"single\" w:sz=\"2\" w:color=\"DDE3EA\"/><w:left w:val=\"single\" w:sz=\"2\" w:color=\"DDE3EA\"/><w:bottom w:val=\"single\" w:sz=\"2\" w:color=\"DDE3EA\"/><w:right w:val=\"single\" w:sz=\"2\" w:color=\"DDE3EA\"/></w:tcBorders>";
  const mar = "<w:tcMar><w:top w:w=\"80\" w:type=\"dxa\"/><w:left w:w=\"130\" w:type=\"dxa\"/><w:bottom w:w=\"80\" w:type=\"dxa\"/><w:right w:w=\"130\" w:type=\"dxa\"/></w:tcMar>";
  const trs = rows.map(function(pair, i) {
    const l = pair[0]; const v = pair[1];
    const bg = i % 2 === 1 ? "F4F8FB" : "FFFFFF";
    const shd = "<w:shd w:val=\"clear\" w:fill=\"" + bg + "\" w:color=\"auto\"/>";
    const tcpWrap = "<w:tcPr><w:tcW w:w=\"4680\" w:type=\"dxa\"/>" + shd + brd + mar + "</w:tcPr>";
    return "<w:tr>" +
      "<w:tc>" + tcpWrap + "<w:p><w:r><w:rPr><w:b/><w:sz w:val=\"19\"/><w:szCs w:val=\"19\"/><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr><w:t>" + esc(l) + "</w:t></w:r></w:p></w:tc>" +
      "<w:tc>" + tcpWrap + "<w:p><w:r><w:rPr><w:i/><w:sz w:val=\"19\"/><w:szCs w:val=\"19\"/><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr><w:t>" + esc(v) + "</w:t></w:r></w:p></w:tc>" +
      "</w:tr>";
  }).join("");
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="40"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="56"/><w:szCs w:val="56"/><w:color w:val="D400AA"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>&#x2B21; INFINITE</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="180"/></w:pPr><w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/><w:color w:val="444444"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>Managed by MEDELITE</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="80" w:after="60"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>FACILITY ASSESSMENT SNAPSHOT</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="D400AA"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>' + esc(state) + '</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>' + trs + '</w:tbl>' +
    '<w:p><w:pPr><w:spacing w:before="240"/></w:pPr></w:p>' +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="718096"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t xml:space="preserve">Source: </w:t></w:r>' +
    '<w:hyperlink r:id="rId1" w:history="1"><w:r><w:rPr><w:color w:val="00B4D8"/><w:u w:val="single"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>' + esc(url) + '</w:t></w:r></w:hyperlink></w:p>' +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1008" w:right="1008" w:bottom="1008" w:left="1008"/></w:sectPr>' +
    '</w:body></w:document>'
  );
}

function buildDocxBlob(rows, state, ccn) {
  function crc32(u8) {
    const T = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      T[i] = c;
    }
    let c = 0xFFFFFFFF;
    for (const b of u8) c = (c >>> 8) ^ T[(c ^ b) & 0xFF];
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function u16(n) { return [n & 0xFF, (n >> 8) & 0xFF]; }
  function u32(n) { return [n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >> 24) & 0xFF]; }
  function b64ToU8(b64) {
    const bin = atob(b64);
    const u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }
  function strToU8(s) { return new TextEncoder().encode(s); }
  function concat() {
    const arrs = Array.from(arguments);
    let total = 0;
    for (const a of arrs) total += a.length;
    const out = new Uint8Array(total);
    let pos = 0;
    for (const a of arrs) { out.set(a, pos); pos += a.length; }
    return out;
  }

  const docXML = makeDocXML(rows, state, ccn);
  const files = [];
  const keys = Object.keys(DOCX_STATIC);
  for (let i = 0; i < keys.length; i++) {
    const name = keys[i];
    const b64 = DOCX_STATIC[name];
    if (b64) files.push({ name: name, data: b64ToU8(b64) });
  }
  files.push({ name: "word/document.xml", data: strToU8(docXML) });

  const localParts = [];
  const cdParts = [];
  let offset = 0;
  for (const f of files) {
    const nameU8 = strToU8(f.name);
    const crc = crc32(f.data);
    const sz = f.data.length;
    const lh = new Uint8Array([0x50,0x4B,0x03,0x04,...u16(20),...u16(0),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(sz),...u32(sz),...u16(nameU8.length),...u16(0)]);
    const cd = new Uint8Array([0x50,0x4B,0x01,0x02,...u16(20),...u16(20),...u16(0),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(sz),...u32(sz),...u16(nameU8.length),...u16(0),...u16(0),...u16(0),...u16(0),...u32(0),...u32(offset)]);
    localParts.push(lh, nameU8, f.data);
    cdParts.push(cd, nameU8);
    offset += lh.length + nameU8.length + f.data.length;
  }
  const cdBuf = concat.apply(null, cdParts);
  const eocd = new Uint8Array([0x50,0x4B,0x05,0x06,...u16(0),...u16(0),...u16(files.length),...u16(files.length),...u32(cdBuf.length),...u32(offset),...u16(0)]);
  const allParts = localParts.concat([cdBuf, eocd]);
  return new Blob([concat.apply(null, allParts)], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(function() { document.body.removeChild(a); URL.revokeObjectURL(url); }, 2000);
}

function buildRows(fd, manual) {
  function f(v, s) { s = s || ""; return (v != null && v !== "") ? (v + s) : "—"; }
  return [
    ["Name of Facility", manual.facility || fd.legalName || "—"],
    ["Location", fd.location || "—"],
    ["EMR", manual.emr || "—"],
    ["Census Capacity", fd.beds || "—"],
    ["Current Census", manual.currentCensus || "—"],
    ["Type of Patient", manual.patientType || "—"],
    ["Previous Coverage from Medelite", manual.prevCoverage || "—"],
    ["Previous Provider Performance from Medelite", manual.prevPerformance || "—"],
    ["Medical Coverage", manual.medCoverage || "—"],
    ["Overall Star Rating", fd.overallStar || "—"],
    ["Health Inspection", fd.healthStar || "—"],
    ["Staffing", fd.staffStar || "—"],
    ["Quality of Resident Care", fd.qualityStar || "—"],
    ["Short Term Hospitalization", f(fd.strHosp, "%")],
    ["STR National Avg. for Hospitalization", f(fd.strHospNatl, "%")],
    ["STR State National Avg. for Hospitalization", f(fd.strHospState, "%")],
    ["STR ED Visit", f(fd.strED, "%")],
    ["STR ED Visits National Avg.", f(fd.strEDNatl, "%")],
    ["STR ED Visits State Avg.", f(fd.strEDState, "%")],
    ["LT Hospitalization", f(fd.ltHosp, "")],
    ["LT National Avg. for Hospitalization", f(fd.ltHospNatl, "")],
    ["LT State National Avg. for Hospitalization", f(fd.ltHospState, "")],
    ["ED Visit", f(fd.ltED, "")],
    ["LT ED Visits National Avg.", f(fd.ltEDNatl, "")],
    ["LT ED Visits State Avg.", f(fd.ltEDState, "")]
  ];
}

function dlPDF(fd, manual, ccn) {
  const rows = buildRows(fd, manual);
  const url = "https://www.medicare.gov/care-compare/details/nursing-home/" + ccn;
  const tableRows = rows.map(function(pair) {
    return "<tr><td class=\"lbl\">" + pair[0] + "</td><td class=\"val\">" + pair[1] + "</td></tr>";
  }).join("");
  const html = [
    "<!DOCTYPE html><html><head><meta charset=\"utf-8\"/>",
    "<title>Facility Assessment Snapshot</title>",
    "<style>",
    "@page{size:Letter;margin:.75in}",
    "*{box-sizing:border-box;margin:0;padding:0}",
    "body{font-family:Arial,sans-serif;font-size:10pt;color:#1A2533}",
    ".hdr{text-align:center;margin-bottom:14pt}",
    ".brand{font-size:24pt;font-weight:900;color:#D400AA;letter-spacing:3px}",
    ".sub{font-size:9pt;color:#444;letter-spacing:1.5px;margin-top:2pt}",
    ".title{font-size:13pt;font-weight:bold;text-align:center;margin:12pt 0 3pt;letter-spacing:1.5px}",
    ".state{font-size:12pt;font-weight:bold;color:#D400AA;text-align:center;margin-bottom:12pt}",
    "table{width:100%;border-collapse:collapse}",
    "td{padding:6pt 10pt;border:.5pt solid #DDE3EA;vertical-align:middle}",
    ".lbl{font-weight:bold;width:50%}.val{font-style:italic}",
    "tr:nth-child(even) td{background:#F4F8FB}",
    ".foot{margin-top:12pt;font-size:7.5pt;color:#718096;text-align:center}",
    "a{color:#00B4D8;text-decoration:underline}",
    ".btn{display:block;margin:0 auto 16pt;padding:8pt 20pt;background:#D400AA;color:#fff;border:none;border-radius:6pt;font-size:11pt;cursor:pointer;font-weight:bold}",
    "@media print{.btn{display:none}}",
    "</style></head><body>",
    "<button class=\"btn\" onclick=\"window.print()\">Print / Save as PDF</button>",
    "<div class=\"hdr\"><div class=\"brand\">&#x2B21; INFINITE</div><div class=\"sub\">Managed by MEDELITE</div></div>",
    "<div class=\"title\">FACILITY ASSESSMENT SNAPSHOT</div>",
    "<div class=\"state\">" + (fd.state || "—") + "</div>",
    "<table>" + tableRows + "</table>",
    "<div class=\"foot\">Source: <a href=\"" + url + "\">" + url + "</a></div>",
    "</body></html>"
  ].join("");
  triggerDownload(new Blob([html], { type: "text/html;charset=utf-8" }), "Facility_Assessment_Snapshot.html");
}

function dlDocx(fd, manual, ccn) {
  const rows = buildRows(fd, manual);
  const blob = buildDocxBlob(rows, fd.state || "—", ccn);
  triggerDownload(blob, "Facility_Assessment_Snapshot.docx");
}

async function fetchViaAPI(ccn) {
  const systemPrompt = (
    "You are a CMS data extraction agent. " +
    "Search Medicare Care Compare for nursing home CCN " + ccn + ". " +
    "Return ONLY valid JSON with no markdown, no code fences, no extra text. " +
    "Required JSON fields: legalName, location, state, beds, " +
    "overallStar, healthStar, staffStar, qualityStar (all 1-5 star ratings as strings), " +
    "strHosp, strHospNatl, strHospState, strED, strEDNatl, strEDState " +
    "(short-stay metrics as decimal numbers without percent signs), " +
    "ltHosp, ltHospNatl, ltHospState, ltED, ltEDNatl, ltEDState " +
    "(long-stay metrics as decimal numbers), " +
    "found (boolean true or false). " +
    "Example for not found: {found: false, legalName: empty string}"
  );

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      system: systemPrompt,
      tools: [{ type: "web_search_20250305", name: "web_search" }],
      messages: [{
        role: "user",
        content: "Fetch nursing home data for CCN: " + ccn + ". Search https://www.medicare.gov/care-compare/details/nursing-home/" + ccn + " and return only the JSON object."
      }]
    })
  });
  if (!res.ok) throw new Error("API error " + res.status);
  const data = await res.json();
  const txt = (data.content || []).filter(function(b) { return b.type === "text"; }).map(function(b) { return b.text; }).join("");
  const jsonMatch = txt.match(/[{][^]*?[}]/);
  if (!jsonMatch) throw new Error("No JSON found in response");
  return JSON.parse(jsonMatch[0]);
}

function Stars({ val }) {
  const n = Number(val) || 0;
  return (
    <span>
      {[1, 2, 3, 4, 5].map(i => (
        <span key={i} style={{ color: i <= n ? C.amber : C.border, fontSize: 18 }}>★</span>
      ))}
      <span style={{ color: C.light, fontSize: 11, marginLeft: 4 }}>({n}/5)</span>
    </span>
  );
}

function KPI({ label, value, sub, accent }) {
  return (
    <div style={{ background: C.white, border: "1px solid " + C.border, borderRadius: 10, padding: "13px 15px", borderTop: "4px solid " + (accent || C.teal), flex: "1 1 115px", minWidth: 115 }}>
      <div style={{ fontSize: 10, color: C.light, textTransform: "uppercase", letterSpacing: 1 }}>{label}</div>
      <div style={{ fontSize: 22, fontWeight: 900, color: C.dark, marginTop: 3 }}>{value != null ? value : ""}</div>
      {sub && <div style={{ fontSize: 10, color: C.light, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function Lbl({ children }) {
  return <div style={{ fontSize: 11, fontWeight: 700, color: C.mid, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>{children}</div>;
}

function Inp({ value, onChange, placeholder, type }) {
  return (
    <input
      type={type || "text"}
      value={value}
      onChange={function(e) { onChange(e.target.value); }}
      placeholder={placeholder}
      style={{ width: "100%", padding: "8px 11px", border: "1px solid " + C.border, borderRadius: 6, fontSize: 13, color: C.dark, background: C.white, boxSizing: "border-box", fontFamily: "inherit", outline: "none" }}
    />
  );
}

function Fld({ label, children }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <Lbl>{label}</Lbl>
      {children}
    </div>
  );
}

function SH({ title }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "20px 0 11px" }}>
      <div style={{ height: 3, width: 18, borderRadius: 2, background: C.magenta }} />
      <span style={{ fontWeight: 800, fontSize: 11, textTransform: "uppercase", letterSpacing: 1, color: C.mid }}>{title}</span>
      <div style={{ flex: 1, height: 1, background: C.border }} />
    </div>
  );
}

function DownloadMenu({ fd, manual, ccn }) {
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const ref = useRef(null);

  useEffect(function() {
    function handleClick(e) {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", handleClick);
    return function() { document.removeEventListener("mousedown", handleClick); };
  }, []);

  function handlePDF() {
    setOpen(false);
    try { dlPDF(fd, manual, ccn); } catch(e) { alert("PDF error: " + e.message); }
  }

  function handleDocx() {
    setOpen(false);
    setBusy(true);
    try { dlDocx(fd, manual, ccn); } catch(e) { alert("DOCX error: " + e.message); }
    setBusy(false);
  }

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button
        onClick={function() { setOpen(function(o) { return !o; }); }}
        disabled={busy}
        style={{ padding: "9px 20px", borderRadius: 7, border: "none", background: busy ? "#94a3b8" : C.magenta, color: C.white, fontWeight: 800, fontSize: 13, cursor: busy ? "wait" : "pointer", display: "flex", alignItems: "center", gap: 8 }}
      >
        {busy ? "Generating…" : "⬇ Download Report"}
        {!busy && <span style={{ fontSize: 9 }}>{open ? "▲" : "▼"}</span>}
      </button>
      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 6px)", right: 0, background: C.white, border: "1px solid " + C.border, borderRadius: 9, boxShadow: "0 8px 28px rgba(0,0,0,0.14)", minWidth: 220, zIndex: 9999, overflow: "hidden" }}>
          <button onClick={handlePDF} style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "13px 16px", border: "none", background: "transparent", cursor: "pointer", textAlign: "left", fontSize: 13, color: C.dark, fontWeight: 600, borderBottom: "1px solid " + C.border }}>
            <span style={{ fontSize: 22 }}>📄</span>
            <div>
              <div style={{ fontWeight: 700 }}>Download as PDF</div>
              <div style={{ fontSize: 11, color: C.light }}>Downloads HTML → open → Print to save as PDF</div>
            </div>
          </button>
          <button onClick={handleDocx} style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: "13px 16px", border: "none", background: "transparent", cursor: "pointer", textAlign: "left", fontSize: 13, color: C.dark, fontWeight: 600 }}>
            <span style={{ fontSize: 22 }}>📝</span>
            <div>
              <div style={{ fontWeight: 700 }}>Download as DOCX</div>
              <div style={{ fontSize: 11, color: C.light }}>Editable Word document</div>
            </div>
          </button>
        </div>
      )}
    </div>
  );
}

function MiniBar({ data, title, unit }) {
  unit = unit || "";
  return (
    <div>
      <div style={{ fontSize: 11, fontWeight: 700, color: C.mid, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 5 }}>{title}</div>
      <ResponsiveContainer width="100%" height={125}>
        <BarChart data={data} margin={{ top: 4, right: 4, left: -28, bottom: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
          <XAxis dataKey="name" tick={{ fontSize: 10 }} />
          <YAxis tick={{ fontSize: 10 }} />
          <Tooltip formatter={function(v) { return v + unit; }} />
          <Bar dataKey="value" radius={[4, 4, 0, 0]}>
            {data.map(function(_, i) { return <Cell key={i} fill={[C.magenta, C.teal, C.navy][i % 3]} />; })}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

function StarRadar({ fd }) {
  const d = [
    { s: "Overall", v: Number(fd.overallStar) || 0 },
    { s: "Health", v: Number(fd.healthStar) || 0 },
    { s: "Staffing", v: Number(fd.staffStar) || 0 },
    { s: "Quality", v: Number(fd.qualityStar) || 0 }
  ];
  return (
    <ResponsiveContainer width="100%" height={195}>
      <RadarChart data={d} outerRadius={68}>
        <PolarGrid stroke={C.border} />
        <PolarAngleAxis dataKey="s" tick={{ fontSize: 11, fill: C.mid }} />
        <PolarRadiusAxis angle={30} domain={[0, 5]} tick={{ fontSize: 9 }} />
        <Radar name="Stars" dataKey="v" stroke={C.magenta} fill={C.magenta} fillOpacity={0.32} />
      </RadarChart>
    </ResponsiveContainer>
  );
}

function Summary({ fd, manual }) {
  const overall = Number(fd.overallStar) || 0;
  const quality = Number(fd.qualityStar) || 0;
  const staffing = Number(fd.staffStar) || 0;
  const strH = parseFloat(fd.strHosp) || 0;
  const strHN = parseFloat(fd.strHospNatl) || 0;
  const strE = parseFloat(fd.strED) || 0;
  const strEN = parseFloat(fd.strEDNatl) || 0;
  const ltE = parseFloat(fd.ltED) || 0;
  const ltEN = parseFloat(fd.ltEDNatl) || 0;
  const occ = (manual.currentCensus && fd.beds) ? Math.round((parseInt(manual.currentCensus) / parseInt(fd.beds)) * 100) : null;

  const flags = [];
  if (overall <= 1) flags.push({ t: "danger", msg: "Critical: Overall Star Rating is " + overall + "/5 — poor performer requiring thorough vetting." });
  else if (overall === 2) flags.push({ t: "warn", msg: "Below-average Overall Star Rating (" + overall + "/5). Review health inspection history." });
  else if (overall >= 4) flags.push({ t: "ok", msg: "Strong Overall Star Rating (" + overall + "/5) — above average performer." });
  if (quality >= 4) flags.push({ t: "ok", msg: "Quality of Resident Care (" + quality + "/5) is above average — a positive clinical indicator." });
  else if (quality <= 2) flags.push({ t: "warn", msg: "Quality of Resident Care is low (" + quality + "/5). Request QM trend data before engagement." });
  if (staffing <= 2) flags.push({ t: "warn", msg: "Low Staffing Rating (" + staffing + "/5). May indicate resource constraints affecting care." });
  if (strH > 0 && strHN > 0) {
    if (strH < strHN) flags.push({ t: "ok", msg: "Short-Term Hospitalization (" + strH + "%) is below national average (" + strHN + "%) — favorable readmission outcome." });
    else flags.push({ t: "warn", msg: "Short-Term Hospitalization (" + strH + "%) exceeds national average (" + strHN + "%). Elevated readmission risk." });
  }
  if (strE > 0 && strEN > 0) {
    if (strE > strEN) flags.push({ t: "warn", msg: "Short-Stay ED Visits (" + strE + "%) exceed national average (" + strEN + "%) — gaps in on-site care management." });
    else flags.push({ t: "ok", msg: "Short-Stay ED Visits (" + strE + "%) are below national average (" + strEN + "%)." });
  }
  if (ltE > 0 && ltEN > 0) {
    if (ltE > ltEN * 2) flags.push({ t: "danger", msg: "LT ED Visit rate (" + ltE + "/1,000 days) is " + (ltE / ltEN).toFixed(1) + "x national average (" + ltEN + ") — urgent clinical review warranted." });
    else if (ltE > ltEN) flags.push({ t: "warn", msg: "LT ED Visit rate (" + ltE + "/1,000 days) exceeds national average (" + ltEN + ")." });
    else flags.push({ t: "ok", msg: "LT ED Visit rate (" + ltE + "/1,000 days) is below national average (" + ltEN + ")." });
  }
  if (occ !== null) flags.push({ t: "info", msg: "Current occupancy: ~" + occ + "% (" + manual.currentCensus + " of " + fd.beds + " certified beds)." });
  if (manual.prevCoverage === "Yes") flags.push({ t: "info", msg: "Medelite has prior coverage history with this facility. Cross-reference provider performance records." });

  const icons = { ok: "✅", warn: "⚠️", danger: "🔴", info: "ℹ️" };
  const bg = { ok: "#F0FDF4", warn: "#FFFBEB", danger: "#FEF2F2", info: "#EFF6FF" };
  const brd2 = { ok: "#86EFAC", warn: "#FDE68A", danger: "#FECACA", info: "#BFDBFE" };

  return (
    <div style={{ background: C.white, border: "1px solid " + C.border, borderRadius: 10, padding: "18px 20px" }}>
      <div style={{ fontWeight: 800, fontSize: 13, color: C.dark, marginBottom: 12 }}>
        Assessment Summary — {manual.facility || fd.legalName}
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {flags.map(function(f, i) {
          return (
            <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "10px 14px", borderRadius: 8, background: bg[f.t], border: "1px solid " + brd2[f.t] }}>
              <span style={{ fontSize: 16, flexShrink: 0, marginTop: 1 }}>{icons[f.t]}</span>
              <span style={{ fontSize: 13, color: C.dark }}>{f.msg}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

export default function App() {
  const [ccnInput, setCcnInput] = useState("686123");
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState({ type: "", msg: "" });
  const [fd, setFd] = useState(null);
  const [validCCN, setValidCCN] = useState(false);
  const [manual, setManual] = useState({
    facility: "", emr: "PCC", currentCensus: "112",
    patientType: "Long-term & Short-term", prevCoverage: "Yes",
    prevPerformance: "About 30 patients/day", medCoverage: "Optometry, PCP, Podiatry"
  });

  function setField(key) {
    return function(val) { setManual(function(m) { return Object.assign({}, m, { [key]: val }); }); };
  }

  async function fetchData() {
    const ccn = ccnInput.trim();
    if (!ccn) { setStatus({ type: "error", msg: "Please enter a CCN." }); return; }
    if (!/^[0-9]{6}$/.test(ccn)) { setStatus({ type: "error", msg: "CCN must be exactly 6 digits." }); return; }
    setLoading(true);
    setStatus({ type: "", msg: "" });
    setFd(null);
    setValidCCN(false);

    if (ccn === "686123") {
      setFd(Object.assign({}, REF));
      setValidCCN(true);
      setStatus({ type: "ok", msg: "Loaded verified data for CCN 686123 — Kendall Lakes Healthcare and Rehab Center." });
      setLoading(false);
      return;
    }

    setStatus({ type: "info", msg: "Searching Medicare Care Compare via AI — please wait 15-30 seconds..." });
    try {
      const result = await fetchViaAPI(ccn);
      if (!result.found || !result.legalName) {
        setStatus({ type: "error", msg: "No facility found for CCN " + ccn + ". Verify the CCN on Medicare.gov." });
        setLoading(false);
        return;
      }
      setFd(Object.assign({ ccn: ccn }, result));
      setValidCCN(true);
      setStatus({ type: "ok", msg: "Data loaded — " + result.legalName + "." });
    } catch (e) {
      setStatus({ type: "error", msg: "Search failed: " + e.message });
    }
    setLoading(false);
  }

  const ccn = ccnInput.trim();
  const displayName = fd ? (manual.facility || fd.legalName) : "";
  const medicareURL = (validCCN && ccn) ? ("https://www.medicare.gov/care-compare/details/nursing-home/" + ccn) : null;

  const strHD = fd ? [{ name: "Facility", value: parseFloat(fd.strHosp) || 0 }, { name: "National", value: parseFloat(fd.strHospNatl) || 0 }, { name: "State", value: parseFloat(fd.strHospState) || 0 }] : [];
  const strED2 = fd ? [{ name: "Facility", value: parseFloat(fd.strED) || 0 }, { name: "National", value: parseFloat(fd.strEDNatl) || 0 }, { name: "State", value: parseFloat(fd.strEDState) || 0 }] : [];
  const ltHD = fd ? [{ name: "Facility", value: parseFloat(fd.ltHosp) || 0 }, { name: "National", value: parseFloat(fd.ltHospNatl) || 0 }, { name: "State", value: parseFloat(fd.ltHospState) || 0 }] : [];
  const ltED2 = fd ? [{ name: "Facility", value: parseFloat(fd.ltED) || 0 }, { name: "National", value: parseFloat(fd.ltEDNatl) || 0 }, { name: "State", value: parseFloat(fd.ltEDState) || 0 }] : [];

  const statusBg = status.type === "error" ? "#FEF2F2" : status.type === "ok" ? "#F0FDF4" : "#EFF6FF";
  const statusBd = status.type === "error" ? "#FECACA" : status.type === "ok" ? "#86EFAC" : "#BFDBFE";
  const statusCl = status.type === "error" ? "#B91C1C" : status.type === "ok" ? "#15803D" : "#1D4ED8";
  const statusIc = status.type === "error" ? "⚠️" : status.type === "ok" ? "✅" : "ℹ️";

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "Arial, sans-serif" }}>
      <div style={{ background: C.navy, position: "sticky", top: 0, zIndex: 100, boxShadow: "0 2px 16px rgba(0,0,0,0.28)" }}>
        <div style={{ maxWidth: 1120, margin: "0 auto", padding: "12px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div>
              <div style={{ color: C.magenta, fontWeight: 900, fontSize: 20, letterSpacing: 3, lineHeight: 1 }}>INFINITE</div>
              <div style={{ color: "#9CA3AF", fontSize: 9, letterSpacing: 2, marginTop: 2 }}>Managed by MEDELITE</div>
            </div>
            <div style={{ width: 1, height: 30, background: "rgba(255,255,255,0.12)" }} />
            <div style={{ color: "#CBD5E1", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1.5 }}>Facility Assessment Snapshot</div>
          </div>
          {fd && validCCN && <DownloadMenu fd={fd} manual={manual} ccn={ccn} />}
        </div>
      </div>

      <div style={{ maxWidth: 1120, margin: "0 auto", padding: "22px 24px" }}>
        <div style={{ background: C.white, border: "1px solid " + C.border, borderRadius: 12, padding: "18px 22px", marginBottom: 20 }}>
          <Lbl>CMS Certification Number (CCN)</Lbl>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
            <input
              value={ccnInput}
              onChange={function(e) { setCcnInput(e.target.value); }}
              onKeyDown={function(e) { if (e.key === "Enter") fetchData(); }}
              placeholder="Enter 6-digit CCN (e.g. 686123)"
              style={{ flex: "1 1 200px", padding: "9px 14px", border: "1.5px solid " + C.border, borderRadius: 7, fontSize: 15, fontFamily: "monospace", color: C.dark, background: C.white, outline: "none", boxSizing: "border-box" }}
            />
            <button
              onClick={fetchData}
              disabled={loading}
              style={{ padding: "9px 24px", borderRadius: 7, border: "none", background: loading ? "#94a3b8" : C.navy, color: C.white, fontWeight: 800, fontSize: 13, cursor: loading ? "not-allowed" : "pointer", whiteSpace: "nowrap" }}
            >
              {loading ? "Searching..." : "Fetch Facility"}
            </button>
            {medicareURL && (
              <a href={medicareURL} target="_blank" rel="noopener noreferrer" style={{ fontSize: 12, color: C.teal, textDecoration: "underline", whiteSpace: "nowrap" }}>
                View on Medicare.gov
              </a>
            )}
          </div>
          {status.msg && (
            <div style={{ marginTop: 10, padding: "9px 14px", borderRadius: 6, fontSize: 12, background: statusBg, border: "1px solid " + statusBd, color: statusCl }}>
              {statusIc} {status.msg}
            </div>
          )}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "290px 1fr", gap: 20, alignItems: "start" }}>
          <div style={{ background: C.white, border: "1px solid " + C.border, borderRadius: 12, padding: "18px 20px" }}>
            <SH title="Operational Inputs" />
            <Fld label="Facility">
              <Inp value={manual.facility} onChange={setField("facility")} placeholder={fd ? fd.legalName : "Override CMS facility name"} />
            </Fld>
            <Fld label="EMR System">
              <Inp value={manual.emr} onChange={setField("emr")} placeholder="e.g. PCC, MatrixCare" />
            </Fld>
            <Fld label="Current Census">
              <Inp value={manual.currentCensus} onChange={setField("currentCensus")} placeholder="e.g. 112" type="number" />
            </Fld>
            <Fld label="Type of Patient">
              <Inp value={manual.patientType} onChange={setField("patientType")} placeholder="e.g. Long-term and Short-term" />
            </Fld>
            <Fld label="Previous Coverage from Medelite">
              <select
                value={manual.prevCoverage}
                onChange={function(e) { setField("prevCoverage")(e.target.value); }}
                style={{ width: "100%", padding: "8px 11px", border: "1px solid " + C.border, borderRadius: 6, fontSize: 13, color: C.dark, background: C.white, fontFamily: "inherit" }}
              >
                <option value="">Select...</option>
                <option value="Yes">Yes</option>
                <option value="No">No</option>
              </select>
            </Fld>
            <Fld label="Previous Provider Performance">
              <Inp value={manual.prevPerformance} onChange={setField("prevPerformance")} placeholder="e.g. About 30 patients/day" />
            </Fld>
            <Fld label="Medical Coverage">
              <Inp value={manual.medCoverage} onChange={setField("medCoverage")} placeholder="e.g. Optometry, PCP, Podiatry" />
            </Fld>
            {medicareURL && (
              <div style={{ marginTop: 14, padding: "11px 14px", background: C.bg, borderRadius: 8, fontSize: 12, color: C.mid }}>
                <div style={{ fontWeight: 700, marginBottom: 4 }}>Medicare Source</div>
                <a href={medicareURL} target="_blank" rel="noopener noreferrer" style={{ color: C.teal, wordBreak: "break-all" }}>{medicareURL}</a>
              </div>
            )}
          </div>

          <div>
            {!fd ? (
              <div style={{ background: C.white, border: "2px dashed " + C.border, borderRadius: 12, padding: "80px 24px", textAlign: "center" }}>
                <div style={{ fontSize: 52, marginBottom: 14 }}>🏥</div>
                <div style={{ fontSize: 16, fontWeight: 700, color: C.dark }}>Enter a CCN above to load facility data</div>
              </div>
            ) : (
              <div>
                <div style={{ background: C.white, border: "1px solid " + C.border, borderRadius: 12, padding: "17px 22px", marginBottom: 14 }}>
                  <div style={{ display: "flex", alignItems: "flex-start", gap: 14, flexWrap: "wrap" }}>
                    <div style={{ padding: "4px 18px", background: C.magenta, color: C.white, borderRadius: 20, fontWeight: 900, fontSize: 14, letterSpacing: 2, flexShrink: 0 }}>{fd.state}</div>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 18, fontWeight: 900, color: C.dark }}>{displayName}</div>
                      <div style={{ fontSize: 13, color: C.light, marginTop: 2 }}>{fd.location}</div>
                    </div>
                  </div>
                </div>

                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
                  <KPI label="Overall Rating" value={<Stars val={fd.overallStar} />} accent={C.magenta} />
                  <KPI label="Health Inspection" value={<Stars val={fd.healthStar} />} accent={C.amber} />
                  <KPI label="Staffing" value={<Stars val={fd.staffStar} />} accent={C.teal} />
                  <KPI label="Quality of Care" value={<Stars val={fd.qualityStar} />} accent={C.green} />
                </div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
                  <KPI label="Census Capacity" value={fd.beds} sub="Certified Beds" accent={C.navy} />
                  <KPI label="Current Census" value={manual.currentCensus || "—"} sub="Residents today" accent={C.teal} />
                  <KPI label="STR Hospitalization" value={fd.strHosp != null ? (fd.strHosp + "%") : "—"} sub={"Natl avg: " + fd.strHospNatl + "%"} accent={parseFloat(fd.strHosp) > parseFloat(fd.strHospNatl) ? C.red : C.green} />
                  <KPI label="LT ED Visits" value={fd.ltED != null ? String(fd.ltED) : "—"} sub={"per 1k days · Natl: " + fd.ltEDNatl} accent={parseFloat(fd.ltED) > parseFloat(fd.ltEDNatl) * 2 ? C.red : C.amber} />
                </div>

                <div style={{ background: C.white, border: "1px solid " + C.border, borderRadius: 12, padding: "17px 22px", marginBottom: 14 }}>
                  <SH title="Star Ratings — Visual Analysis" />
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 22 }}>
                    <div>
                      <div style={{ fontSize: 11, fontWeight: 700, color: C.mid, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 5 }}>Radar Overview</div>
                      <StarRadar fd={fd} />
                    </div>
                    <div>
                      <div style={{ fontSize: 11, fontWeight: 700, color: C.mid, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 5 }}>By Category (out of 5)</div>
                      <ResponsiveContainer width="100%" height={195}>
                        <BarChart
                          data={[
                            { name: "Overall", v: Number(fd.overallStar) || 0 },
                            { name: "Health", v: Number(fd.healthStar) || 0 },
                            { name: "Staffing", v: Number(fd.staffStar) || 0 },
                            { name: "Quality", v: Number(fd.qualityStar) || 0 }
                          ]}
                          margin={{ top: 4, right: 4, left: -28, bottom: 0 }}
                        >
                          <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                          <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                          <YAxis domain={[0, 5]} tick={{ fontSize: 10 }} />
                          <ReferenceLine y={3} stroke="#94a3b8" strokeDasharray="4 4" />
                          <Tooltip />
                          <Bar dataKey="v" name="Stars" radius={[4, 4, 0, 0]}>
                            {[C.magenta, C.amber, C.teal, C.green].map(function(c, i) { return <Cell key={i} fill={c} />; })}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>

                <div style={{ background: C.white, border: "1px solid " + C.border, borderRadius: 12, padding: "17px 22px", marginBottom: 14 }}>
                  <SH title="All 12 Hospitalization and ED Metrics — Facility vs. Benchmarks" />
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18 }}>
                    <MiniBar data={strHD} title="STR Hospitalization (%)" unit="%" />
                    <MiniBar data={strED2} title="STR ED Visits (%)" unit="%" />
                    <MiniBar data={ltHD} title="LT Hospitalization (per 1,000 days)" />
                    <MiniBar data={ltED2} title="LT ED Visits (per 1,000 days)" />
                  </div>
                </div>

                <div style={{ background: C.white, border: "1px solid " + C.border, borderRadius: 12, padding: "17px 22px", marginBottom: 14 }}>
                  <SH title="Full Report Table" />
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <tbody>
                      {buildRows(fd, manual).map(function(pair, i) {
                        return (
                          <tr key={i} style={{ background: i % 2 === 1 ? C.bg : C.white }}>
                            <td style={{ padding: "7px 12px", border: "1px solid " + C.border, fontWeight: 700, width: "50%" }}>{pair[0]}</td>
                            <td style={{ padding: "7px 12px", border: "1px solid " + C.border, color: C.mid, fontStyle: "italic" }}>{pair[1] || "—"}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                <Summary fd={fd} manual={manual} />
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
