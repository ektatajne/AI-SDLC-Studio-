import React, { useState, useEffect } from 'react';
import { 
  ShieldCheck, 
  CheckCircle2, 
  AlertCircle, 
  RefreshCw, 
  Check,
  X,
  FileDown
} from 'lucide-react';
import { apiFetch, API_BASE, API_KEY } from '../services/api';

interface TestingPanelProps {
  projectId: string;
  projectDetails: any;
  fetchProjectDetails: (id: string) => Promise<void>;
  setToastMessage: (msg: string | null) => void;
}

export default function TestingPanel({
  projectId,
  projectDetails,
  fetchProjectDetails,
  setToastMessage
}: TestingPanelProps) {
  const [loading, setLoading] = useState(false);
  const [testingResult, setTestingResult] = useState<any>(null);
  const [executing, setExecuting] = useState(false);
  const [submittingReview, setSubmittingReview] = useState(false);
  const [reviewerName, setReviewerName] = useState('Lead Tester');
  const [reviewComments, setReviewComments] = useState('');

  const loadTestingResult = async () => {
    setLoading(true);
    try {
      const res = await apiFetch(`${API_BASE}/projects/${projectId}/testing/result`);
      if (res.ok) {
        const data = await res.json();
        setTestingResult(data);
      }
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadTestingResult();
  }, [projectId]);

  const handleStartTesting = async () => {
    setExecuting(true);
    setToastMessage("Starting Testing Phase... this may take a few minutes.");
    try {
      const res = await apiFetch(`${API_BASE}/projects/${projectId}/testing/start`, {
        method: 'POST'
      });
      if (res.ok) {
        const data = await res.json();
        setTestingResult(data);
        setToastMessage("Testing completed successfully.");
        await fetchProjectDetails(projectId);
      } else {
        alert("Testing failed.");
      }
    } catch (err) {
      console.error(err);
      alert("Error starting tests.");
    } finally {
      setExecuting(false);
    }
  };

  const handleReview = async (status: 'APPROVED' | 'REJECTED') => {
    setSubmittingReview(true);
    try {
      const res = await apiFetch(`${API_BASE}/projects/${projectId}/testing/approve`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          status,
          comments: reviewComments,
          reviewer_name: reviewerName
        })
      });
      
      if (res.ok) {
        setToastMessage(`Testing phase ${status.toLowerCase()} successfully.`);
        await fetchProjectDetails(projectId);
      } else {
        const error = await res.json();
        alert(`Failed to submit review: ${error.detail || 'Unknown error'}`);
      }
    } catch (err) {
      console.error(err);
      alert("Error submitting review.");
    } finally {
      setSubmittingReview(false);
    }
  };
  
  const downloadReport = () => {
    const url = `${API_BASE}/projects/${projectId}/testing/report/download/pdf?api_key=${API_KEY}`;
    window.open(url, '_blank');
  };

  const hasResult = testingResult && testingResult.execution_summary;
  const summary = testingResult?.execution_summary;
  const isPending = projectDetails.project.status === 'AWAITING_APPROVAL' && projectDetails.project.current_phase === 'TESTING';

  return (
    <div className="flex-1 flex flex-col min-w-0 bg-[#0b0f19] p-8 overflow-y-auto space-y-6">
      
      <div className="flex items-center justify-between border-b border-slate-800 pb-4">
        <h2 className="text-xl font-black text-slate-200 flex items-center gap-3">
          <ShieldCheck className="h-6 w-6 text-indigo-400" />
          Testing Agent Dashboard
        </h2>
        
        {hasResult && (
          <button 
            onClick={downloadReport}
            className="flex items-center gap-2 py-2 px-4 bg-indigo-600/20 hover:bg-indigo-600/40 text-indigo-300 rounded border border-indigo-500/30 font-bold text-xs"
          >
            <FileDown className="h-4 w-4" /> Download Report
          </button>
        )}
      </div>

      {!hasResult && !executing && (
        <div className="bg-[#111827] border border-slate-800 rounded-xl p-8 text-center space-y-4">
          <ShieldCheck className="h-12 w-12 text-slate-600 mx-auto" />
          <h3 className="text-lg font-bold text-slate-300">Testing Phase Ready</h3>
          <p className="text-sm text-slate-400 max-w-md mx-auto">
            The development phase is complete. You can now invoke the autonomous Testing Agent to generate and execute comprehensive test suites against the finalized codebase.
          </p>
          <button
            onClick={handleStartTesting}
            className="mt-4 flex items-center gap-2 py-2 px-6 bg-indigo-600 hover:bg-indigo-500 text-white font-extrabold rounded-lg mx-auto"
          >
            Start Testing Workflow
          </button>
        </div>
      )}
      
      {executing && (
        <div className="bg-[#111827] border border-slate-800 rounded-xl p-8 text-center space-y-4">
          <RefreshCw className="h-10 w-10 text-indigo-400 animate-spin mx-auto" />
          <h3 className="text-lg font-bold text-slate-300">Testing Agent Running</h3>
          <p className="text-sm text-slate-400 max-w-md mx-auto">
            The Testing Agent is analyzing the source code, generating test cases, and executing them. This might take a few minutes.
          </p>
        </div>
      )}

      {hasResult && !executing && (
        <div className="space-y-6">
          <div className="grid grid-cols-5 gap-4">
             <div className="bg-[#111827] border border-slate-800 rounded-xl p-4 shadow">
               <span className="text-[10px] text-slate-500 font-bold uppercase block">Total Tests</span>
               <span className="text-2xl font-black text-slate-200">{summary?.total || 0}</span>
             </div>
             <div className="bg-[#111827] border border-slate-800 rounded-xl p-4 shadow">
               <span className="text-[10px] text-emerald-500 font-bold uppercase block">Passed</span>
               <span className="text-2xl font-black text-emerald-400">{summary?.passed || 0}</span>
             </div>
             <div className="bg-[#111827] border border-slate-800 rounded-xl p-4 shadow">
               <span className="text-[10px] text-rose-500 font-bold uppercase block">Failed</span>
               <span className="text-2xl font-black text-rose-400">{summary?.failed || 0}</span>
             </div>
             <div className="bg-[#111827] border border-slate-800 rounded-xl p-4 shadow">
               <span className="text-[10px] text-amber-500 font-bold uppercase block">Errors</span>
               <span className="text-2xl font-black text-amber-400">{summary?.errors || 0}</span>
             </div>
             <div className="bg-[#111827] border border-slate-800 rounded-xl p-4 shadow">
               <span className="text-[10px] text-slate-400 font-bold uppercase block">Skipped</span>
               <span className="text-2xl font-black text-slate-300">{summary?.skipped || 0}</span>
             </div>
          </div>

          <div className="bg-[#111827] border border-slate-800 rounded-xl p-6">
             <h3 className="text-sm font-bold text-slate-300 mb-4 uppercase tracking-wider">Quality Gate</h3>
             <div className="flex gap-4 items-center">
               <span className={`px-4 py-2 rounded-lg font-bold ${
                 testingResult?.quality_gate?.overall_status === 'PASS' ? 'bg-emerald-500/20 text-emerald-400' : 
                 testingResult?.quality_gate?.overall_status === 'FAIL' ? 'bg-rose-500/20 text-rose-400' : 
                 'bg-amber-500/20 text-amber-400'
               }`}>
                 STATUS: {testingResult?.quality_gate?.overall_status || 'UNKNOWN'}
               </span>
               <span className="text-sm font-semibold text-slate-400">
                 Release Readiness: {testingResult?.quality_gate?.release_readiness || 'N/A'}
               </span>
               <span className="text-sm font-semibold text-slate-400">
                 Execution Status: {testingResult?.execution_status || 'completed'}
               </span>
             </div>
          </div>

          {isPending && (
             <div className="bg-[#121625] border border-indigo-500/30 p-6 rounded-xl space-y-4">
               <h3 className="text-sm font-bold text-indigo-400 mb-2 uppercase tracking-wider">Testing Review Sign-off</h3>
               <div className="grid grid-cols-2 gap-4">
                 <div>
                   <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Reviewer</label>
                   <input 
                     value={reviewerName}
                     onChange={(e) => setReviewerName(e.target.value)}
                     className="w-full bg-slate-900 border border-slate-700 rounded-lg p-2 text-sm text-slate-200"
                   />
                 </div>
                 <div>
                   <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Comments</label>
                   <input 
                     value={reviewComments}
                     onChange={(e) => setReviewComments(e.target.value)}
                     className="w-full bg-slate-900 border border-slate-700 rounded-lg p-2 text-sm text-slate-200"
                   />
                 </div>
               </div>
               <div className="flex gap-4 pt-2">
                 <button 
                   onClick={() => handleReview('APPROVED')}
                   disabled={submittingReview}
                   className="flex-1 bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-2 rounded-lg transition"
                 >
                   Approve & Complete
                 </button>
                 <button 
                   onClick={() => handleReview('REJECTED')}
                   disabled={submittingReview}
                   className="flex-1 bg-rose-600/20 hover:bg-rose-600/30 text-rose-400 border border-rose-500/30 font-bold py-2 rounded-lg transition"
                 >
                   Reject
                 </button>
               </div>
             </div>
          )}
        </div>
      )}
    </div>
  );
}
