import { useId } from "react"

export default interface TimeProps {
  timestamp: Date
}

export default function TimeProps({ timestamp }: TimeProps) {
  const id = useId();

  return (
    <>
      <button commandfor={id} command="show-modal">{timestamp.toLocaleTimeString()}</button>
      <dialog id={id}></dialog>
    </>
	)
}
