import { useEffect, useState } from "react";
import DocumentEditor from "./DocumentEditor";
import StartScreen from "./StartScreen";
import { routeDocumentId } from "./windows";
import "./App.css";

export default function App() {
  const [documentId, setDocumentId] = useState(routeDocumentId);
  useEffect(() => {
    const update = () => setDocumentId(routeDocumentId());
    addEventListener("popstate", update);
    return () => removeEventListener("popstate", update);
  }, []);
  return documentId ? <DocumentEditor key={documentId} documentId={documentId} /> : <StartScreen />;
}
