import { Component, type ErrorInfo, type PropsWithChildren, type ReactNode } from "react";
import { ErrorState } from "./ErrorState";

interface ErrorBoundaryProps extends PropsWithChildren {
  onOpenAnother: () => void;
}

interface ErrorBoundaryState {
  error?: Error;
}

export class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  override state: ErrorBoundaryState = {};

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error };
  }

  override componentDidCatch(error: Error, info: ErrorInfo): void {
    if (import.meta.env.DEV) {
      console.error("Viewer crashed", error, info);
    }
  }

  override render(): ReactNode {
    if (this.state.error) {
      return <ErrorState error={this.state.error} onOpenAnother={this.props.onOpenAnother} />;
    }

    return this.props.children;
  }
}
