import {
  Modal,
  ModalOverlay,
  ModalContent,
  ModalHeader,
  ModalFooter,
  ModalBody,
  ModalCloseButton,
  useDisclosure,
  Button,
  Text,
  Flex,
  Box,
} from "@chakra-ui/react";
import { useEffect, useState } from "react";
import Calendar from "react-calendar";

interface SummaryModelProps {
  data: {
    id: string;
    name: string;
    duration: string;
    skills: string;
    status: string;
    test: string;
  };
}

function SummaryModel({ data }: SummaryModelProps): JSX.Element {
  const { isOpen, onOpen, onClose } = useDisclosure();
  const [markedDates, setMarkedDates] = useState<Date[]>([]);

  const handleDownloadCSV = () => {
    const csvContent =
      "data:text/csv;charset=utf-8," +
      "Id,Name,Duration,Skills,Status,test\n" +
      (
        [
          "id",
          "name",
          "duration",
          "skills",
          "status",
          "test",
        ] as (keyof typeof data)[]
      )
        .map((key) => data[key])
        .join(",");

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "summary-report.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const markDates = () => {
    const today = new Date();
    const duration = parseInt(data.duration, 10);
  console.log("Duration:", duration);
  const durationMark = duration -1;

  if (isNaN(duration)) {
    console.error("Invalid duration:", durationMark);
    return;
  }

    const endDate = new Date(today);
    endDate.setDate(today.getDate() + durationMark);

    const datesToMark = [];
    while (today <= endDate) {
      datesToMark.push(new Date(today));
      today.setDate(today.getDate() + 1);
    }
    console.log("Marked Dates:", datesToMark);
    setMarkedDates(datesToMark);
  };
  

  useEffect(() => {
    markDates();
  }, [isOpen]);
  return (
    <>
      <Button onClick={onOpen} colorScheme="teal">
        Calender view
      </Button>

      <Modal isOpen={isOpen} onClose={onClose} size="lg">
        <ModalOverlay />
        <ModalContent>
          <ModalHeader bg={"#7928CA"} borderRadius={"4px"} color={"white"}>Summary Details</ModalHeader>
          <ModalCloseButton />
          <ModalBody>
            <Box>
              {/* Left side */}
              <Box >
                <Flex alignItems="center" mb={2}>
                  <Text fontWeight="bold" width="80px" color={"black"}>
                    ID:
                  </Text >{" "}
                  <Text color={"#7928CA"} >
                  {data.id}
                  </Text>
                </Flex>
                <Flex alignItems="center" mb={2}>
                  <Text fontWeight="bold" width="80px" color={"black"}>
                    Name:
                  </Text>{" "}
                  <Text color={"#7928CA"} >
                  {data.name}
                  </Text>
                </Flex>
                <Flex alignItems="center" mb={2}>
                  <Text fontWeight="bold" width="80px" color={"black"}>
                    Duration:
                  </Text>{" "}
                  <Text color={"#7928CA"} >
                  {data.duration}
                  </Text>
                </Flex>
                <Flex alignItems="center" mb={2}>
                  <Text fontWeight="bold" width="80px" color={"black"}>
                    Skills:
                  </Text>{" "}
                  <Text color={"#7928CA"} >
                  {data.skills}
                  </Text>
                </Flex>
                <Flex alignItems="center" mb={2}>
                  <Text fontWeight="bold" width="80px" color={"black"}>
                    Status:
                  </Text>{" "}
                  <Text color={"#7928CA"} >
                  {data.status}
                  </Text>
                </Flex>
                <Flex alignItems="center" mb={2}>
                  <Text fontWeight="bold" width="80px" color={"black"}>
                    Test:
                  </Text>{" "}
                  <Text color={"#7928CA"} >
                  {data.test}
                  </Text>
                </Flex>
              </Box>

              <Box bgGradient='linear(to-l, #7928CA, #7928CA)' borderRadius={"7px"} p={"13px"} boxShadow="xl" // Add box shadow for depth
              transform="rotateX(10deg)" // Tilt the card
              transition="transform 0.3s ease-in-out" // Add transition for smooth effect
              _hover={{
                transform: "rotateX(0deg)", // Reset transform on hover
              }}>
                <Calendar
                  className="custom-calendar"
                  tileClassName={({ date }) =>
                    markedDates.some(
                      (d) => d.toDateString() === date.toDateString()
                    )
                      ? "green-background"
                      : ""
                  }
                  
                />
                  <style>{`
                  .green-background {
                    background-color: #90EE90 !important;
                  }
                  .custom-calendar{
                    font-size:17px;
                    color: white;
                  }
                `}</style>
              </Box>

            </Box>
          </ModalBody>

          <ModalFooter>
            <Button variant="ghost" onClick={onClose}>
              Close
            </Button>
            <Button
              colorScheme="blue"
              variant="ghost"
              onClick={handleDownloadCSV}
            >
              Download CSV
            </Button>
          </ModalFooter>
        </ModalContent>
      </Modal>
    </>
  );
}
export default SummaryModel;
