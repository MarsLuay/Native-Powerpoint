import { Component, createRef, type ErrorInfo, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { errorLog } from './logger';
import { DocxReactView, type DocxReactViewHandle, type DocxReactViewProps } from './DocxReactView';

export interface DocxReactMount {
	getHandle: () => DocxReactViewHandle | null;
	render: (props: DocxReactViewProps) => void;
	unmount: () => void;
}

interface DocxReactErrorBoundaryProps {
	children: ReactNode;
	fileName: string;
}

interface DocxReactErrorBoundaryState {
	error: Error | null;
}

class DocxReactErrorBoundary extends Component<DocxReactErrorBoundaryProps, DocxReactErrorBoundaryState> {
	state: DocxReactErrorBoundaryState = {
		error: null,
	};

	static getDerivedStateFromError(error: Error): DocxReactErrorBoundaryState {
		return { error };
	}

	componentDidCatch(error: Error, errorInfo: ErrorInfo) {
		errorLog('render', `React render failed for ${this.props.fileName}`, {
			error,
			componentStack: errorInfo.componentStack,
		});
	}

	componentDidUpdate(previousProps: DocxReactErrorBoundaryProps) {
		if (previousProps.fileName !== this.props.fileName && this.state.error) {
			this.setState({ error: null });
		}
	}

	render() {
		if (this.state.error) {
			return (
				<div className="docxidian-editor-load-error">
					{`Could not render DOCX editor: ${this.state.error.message}`}
				</div>
			);
		}

		return this.props.children;
	}
}

export function createDocxReactMount(hostEl: HTMLElement): DocxReactMount {
	const ref = createRef<DocxReactViewHandle>();
	const root: Root = createRoot(hostEl);

	return {
		getHandle: () => ref.current,
		render: (props) => {
			root.render(
				<DocxReactErrorBoundary fileName={props.file?.name ?? 'DOCX'}>
					<DocxReactView ref={ref} {...props} />
				</DocxReactErrorBoundary>,
			);
		},
		unmount: () => root.unmount(),
	};
}
