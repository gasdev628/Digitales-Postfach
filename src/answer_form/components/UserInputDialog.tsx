import { Button, Dialog, DialogContent, Textarea } from "@fluentui/react-components";
// import { Pencil } from "lucide-react";
// import { Button } from "@/components/ui/button";
// import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
// import { Textarea } from "@/components/ui/textarea";
import { useState } from "react";
import React from "react";

interface UserInputDialogProps {
  isOpen: boolean;
  onClose: () => void;
  onSubmit: (input: string) => void;
  initialValue?: string;
}

export default function UserInputDialog({ isOpen, onClose, onSubmit, initialValue = "" }: UserInputDialogProps) {
  const [input, setInput] = useState(initialValue);

  const handleSubmit = () => {
    onSubmit(input);
    onClose();
  };

  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent className="sm:max-w-[600px]">
        {/* <DialogHeader className="flex flex-row items-center space-y-0 space-x-2">
          <Pencil className="h-5 w-5 text-gray-500" />
          <DialogTitle>User Input</DialogTitle>
        </DialogHeader> */}

        <div className="grid gap-4 py-4">
          <Textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="Enter your text here..."
            className="min-h-[200px] resize-none"
          />
        </div>

        <div className="flex justify-end space-x-2">
          <Button onClick={onClose} className="bg-gray-200 hover:bg-gray-300 text-gray-800">
            Cancel
          </Button>
          <Button onClick={handleSubmit} className="bg-[#166F5A] hover:bg-[#166F5A]/90 text-white">
            OK
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}

// export default UserInputDialog;
