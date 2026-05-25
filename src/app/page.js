'use client';
import dynamic from 'next/dynamic';

const ChemSegTool = dynamic(() => import('@/components/ChemSegTool'), { ssr: false });

export default function Home() {
  return <ChemSegTool />;
}
