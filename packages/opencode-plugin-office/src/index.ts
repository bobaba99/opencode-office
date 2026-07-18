import type { Plugin } from "@opencode-ai/plugin"
import { officeTools } from "./tools"

export const OfficePlugin: Plugin = async () => ({ tool: officeTools })
export default OfficePlugin
